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

    // ---------- StyleRewriter 样式查找兜底（#116）----------

    /// <summary>
    /// 有些文档把正文样式的 styleId 本地化成 a / 1 之类，w:name 仍是 Normal。
    /// 只按 styleId 精确匹配会认定「没有正文样式」，整条格式化链断掉。
    /// </summary>
    [Fact]
    public void RewriterFallsBackToStyleNameWhenStyleIdIsLocalized()
    {
        var path = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"), P("正文段落", style: "a"));
            CreateDocx(path, body, ("a", "Normal"));

            Assert.True(new StyleRewriter().HasStyle(path, "Normal"));

            new StyleRewriter().SetStyleFontSize(path, path, "Normal", 12);

            var style = ReadStyle(path, "a");
            Assert.Equal("24", style.Element(W("rPr"))?.Element(W("sz"))?.Attribute(W("val"))?.Value);
        }
        finally { TryDelete(path); }
    }

    /// <summary>
    /// Word 内置名带空格（heading 1），请求方习惯写 Heading1，要能对上。
    /// </summary>
    [Fact]
    public void RewriterMatchesBuiltInNameIgnoringSpacesAndCase()
    {
        var path = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"), P("一级标题", style: "x9"));
            CreateDocx(path, body, ("x9", "heading 1"));

            new StyleRewriter().SetStyleFontSize(path, path, "Heading1", 16);

            var style = ReadStyle(path, "x9");
            Assert.Equal("32", style.Element(W("rPr"))?.Element(W("sz"))?.Attribute(W("val"))?.Value);
        }
        finally { TryDelete(path); }
    }

    [Fact]
    public void HasStyleReturnsFalseWhenNeitherIdNorNameMatches()
    {
        var path = NewTempDocxPath();
        try
        {
            CreateDocx(path, new XElement(W("body"), P("正文", style: "Normal")), ("Normal", "Normal"));

            Assert.True(new StyleRewriter().HasStyle(path, "Normal"));
            Assert.False(new StyleRewriter().HasStyle(path, "GhostStyle"));
        }
        finally { TryDelete(path); }
    }

    /// <summary>
    /// 样式表里没有 Heading1…Caption 时，full-format 要跳过并说明，而不是整条命令中止。
    /// </summary>
    [Fact]
    public void FullFormatSkipsMissingStylesInsteadOfAborting()
    {
        var path = NewTempDocxPath();
        var outputPath = NewTempDocxPath();
        var backupDir = Path.Combine(Path.GetTempPath(), "nku-undo-tests-" + Guid.NewGuid().ToString("N"));
        try
        {
            CreateDocx(path, new XElement(W("body"), P("正文段落", style: "a")), ("a", "Normal"));

            var output = new StringWriter();
            var error = new StringWriter();
            var command = NewFullFormatCommand(CliAssembly, output, error, backupDir);

            var exitCode = Execute(command, new[] { path, outputPath });

            Assert.Equal(0, exitCode);
            Assert.True(File.Exists(outputPath));
            // 正文样式走名字兜底认出来了，不该出现在跳过名单里
            Assert.Contains("跳过", output.ToString());
            Assert.Contains("Heading1", output.ToString());
            Assert.DoesNotContain("没有 Normal", output.ToString());
        }
        finally
        {
            TryDelete(path);
            TryDelete(outputPath);
            try { Directory.Delete(backupDir, true); } catch (IOException) { }
        }
    }

    [Fact]
    public void AssignerResolvesStyleByWNameWhenStyleIdIsLocalized()
    {
        var path = NewTempDocxPath();
        var output = NewTempDocxPath();
        try
        {
            // 真实文档里常见：正文样式的 styleId 是 a，w:name 才是 Normal
            CreateDocx(path, new XElement(W("body"), P("本章介绍研究背景。")), ("a", "Normal"));

            var writer = new StringWriter();
            var result = new StyleAutoAssigner(writer).AutoAssignStyles(path, output);

            Assert.Equal(1, result.TotalAssigned);
            Assert.Equal(0, result.ParagraphsWithoutUsableStyle);
            // 写进去的必须是样式表里真实存在的 id，不能硬写 "Normal"
            Assert.Equal("a", ReadAssignedStyles(output)["本章介绍研究背景。"]);
        }
        finally { TryDelete(path); TryDelete(output); }
    }

    [Fact]
    public void AssignerReportsParagraphsWithoutUsableStyle()
    {
        var path = NewTempDocxPath();
        var output = NewTempDocxPath();
        try
        {
            // 样式表里连正文样式都没有，套不上就得说出来，不能报「无需调整」
            CreateDocx(path, new XElement(W("body"), P("一段正文。")), ("X", "Body Text"));

            var result = new StyleAutoAssigner(new StringWriter()).AutoAssignStyles(path, output);

            Assert.Equal(0, result.TotalAssigned);
            Assert.Equal(1, result.ParagraphsWithoutUsableStyle);
        }
        finally { TryDelete(path); TryDelete(output); }
    }

    [Fact]
    public void AnalyzerIgnoresParagraphMarkFormatting()
    {
        var path = NewTempDocxPath();
        try
        {
            // 段落标记上的加粗（pPr/rPr）只作用于回车符，不影响正文，不该报成「被压住」
            var para = new XElement(W("p"),
                new XElement(W("pPr"),
                    new XElement(W("pStyle"), new XAttribute(W("val"), "Normal")),
                    new XElement(W("rPr"), new XElement(W("b")))),
                new XElement(W("r"), new XElement(W("t"), "段落标记带加粗的正文")));

            CreateDocx(path, new XElement(W("body"), para), ("Normal", "Normal"));

            var result = new StyleAnalyzer().Analyze(path);

            Assert.Empty(result.DirectFormatOverrides);
        }
        finally { TryDelete(path); }
    }

    // ---------- Issue #118：清除压过样式的直接格式 ----------
    [Fact]
    public void CleanerRemovesOverridesButKeepsSuperscriptAndLanguage()
    {
        var path = NewTempDocxPath();
        try
        {
            var run = new XElement(W("r"),
                new XElement(W("rPr"),
                    new XElement(W("rFonts"),
                        new XAttribute(W("ascii"), "Times New Roman"),
                        new XAttribute(W("eastAsia"), "宋体")),
                    new XElement(W("sz"), new XAttribute(W("val"), "21")),
                    new XElement(W("szCs"), new XAttribute(W("val"), "21")),
                    new XElement(W("b")),
                    new XElement(W("vertAlign"), new XAttribute(W("val"), "superscript")),
                    new XElement(W("lang"), new XAttribute(W("val"), "en-US"))),
                new XElement(W("t"), "上标引用"));

            var para = new XElement(W("p"),
                new XElement(W("pPr"), new XElement(W("pStyle"), new XAttribute(W("val"), "Normal"))),
                run);

            CreateDocx(path, new XElement(W("body"), para), ("Normal", "Normal"));

            var result = new DirectFormatCleaner().Clean(path, path);

            Assert.Equal(1, result.RunsTouched);

            var rPr = FirstRunProperties(path);
            Assert.NotNull(rPr);
            // 该由样式管的，全清掉
            Assert.Null(rPr!.Element(W("rFonts")));
            Assert.Null(rPr.Element(W("sz")));
            Assert.Null(rPr.Element(W("szCs")));
            Assert.Null(rPr.Element(W("b")));
            // 不属于「格式与样式打架」的，一个都不能丢
            Assert.NotNull(rPr.Element(W("vertAlign")));
            Assert.NotNull(rPr.Element(W("lang")));
        }
        finally { TryDelete(path); }
    }

    [Fact]
    public void CleanerSkipsParagraphsWithoutUsableStyle()
    {
        var path = NewTempDocxPath();
        try
        {
            var para = new XElement(W("p"),
                new XElement(W("r"),
                    new XElement(W("rPr"), new XElement(W("sz"), new XAttribute(W("val"), "21"))),
                    new XElement(W("t"), "没有样式的段落")));

            CreateDocx(path, new XElement(W("body"), para), ("Normal", "Normal"));

            var result = new DirectFormatCleaner().Clean(path, path);

            // 没有样式时直接格式就是它唯一的格式，清掉等于把版面抹平
            Assert.Equal(1, result.ParagraphsSkippedNoStyle);
            Assert.Equal(0, result.ParagraphsTouched);
            Assert.NotNull(FirstRunProperties(path)?.Element(W("sz")));
        }
        finally { TryDelete(path); }
    }

    [Fact]
    public void CleanerKeepsParagraphSpacingAndAlignment()
    {
        var path = NewTempDocxPath();
        try
        {
            var para = new XElement(W("p"),
                new XElement(W("pPr"),
                    new XElement(W("pStyle"), new XAttribute(W("val"), "Normal")),
                    new XElement(W("jc"), new XAttribute(W("val"), "center")),
                    new XElement(W("ind"), new XAttribute(W("firstLineChars"), "200")),
                    new XElement(W("spacing"),
                        new XAttribute(W("before"), "240"),
                        new XAttribute(W("after"), "240"),
                        new XAttribute(W("line"), "360"),
                        new XAttribute(W("lineRule"), "auto"))),
                new XElement(W("r"), new XElement(W("t"), "居中的图题")));

            CreateDocx(path, new XElement(W("body"), para), ("Normal", "Normal"));

            new DirectFormatCleaner().Clean(path, path);

            var pPr = ReadDocumentXml(path).Root!.Element(W("body"))!.Element(W("p"))!.Element(W("pPr"))!;

            // 缩进和行距归样式管
            Assert.Null(pPr.Element(W("ind")));
            var spacing = pPr.Element(W("spacing"))!;
            Assert.Null(spacing.Attribute(W("line")));
            Assert.Null(spacing.Attribute(W("lineRule")));
            // 段前段后和对齐保留：标题靠它们撑版面，样式里没有对应定义
            Assert.Equal("240", spacing.Attribute(W("before"))!.Value);
            Assert.Equal("240", spacing.Attribute(W("after"))!.Value);
            Assert.Equal("center", pPr.Element(W("jc"))!.Attribute(W("val"))!.Value);
        }
        finally { TryDelete(path); }
    }

    [Fact]
    public void AnalyzerReportsStyledParagraphStillOverriddenByDirectFormat()
    {
        var path = NewTempDocxPath();
        try
        {
            var overridden = new XElement(W("p"),
                new XElement(W("pPr"), new XElement(W("pStyle"), new XAttribute(W("val"), "Normal"))),
                new XElement(W("r"),
                    new XElement(W("rPr"), new XElement(W("sz"), new XAttribute(W("val"), "21"))),
                    new XElement(W("t"), "套了样式但字号是手动的")));

            CreateDocx(path, new XElement(W("body"), overridden, P("干净段落", style: "Normal")),
                ("Normal", "Normal"));

            var result = new StyleAnalyzer().Analyze(path);

            // 以前这里报 0 条：套了样式就算「已归属」，直接格式完全看不见
            var item = Assert.Single(result.DirectFormatOverrides);
            Assert.Contains("手动字号", item.Overrides);
            Assert.Contains("套了样式", item.Preview);
        }
        finally { TryDelete(path); }
    }

    [Fact]
    public void FullFormatClearsDirectFormattingSoStyleTakesEffect()
    {
        var path = NewTempDocxPath();
        var outputPath = NewTempDocxPath();
        var backupDir = Path.Combine(Path.GetTempPath(), "nku-undo-tests-" + Guid.NewGuid().ToString("N"));
        try
        {
            var para = new XElement(W("p"),
                new XElement(W("pPr"), new XElement(W("pStyle"), new XAttribute(W("val"), "Normal"))),
                new XElement(W("r"),
                    new XElement(W("rPr"), new XElement(W("sz"), new XAttribute(W("val"), "21"))),
                    new XElement(W("t"), "字号被手动设成五号的正文")));

            CreateDocx(path, new XElement(W("body"), para), ("Normal", "Normal"));

            var output = new StringWriter();
            var error = new StringWriter();
            var command = NewFullFormatCommand(CliAssembly, output, error, backupDir);

            var exitCode = Execute(command, new[] { path, outputPath });

            Assert.Equal(0, exitCode);
            Assert.Contains("【2/6】清除压过样式的手动格式", output.ToString());

            // run 上的手动字号没了，样式才真正说了算
            Assert.Null(FirstRunProperties(outputPath)?.Element(W("sz")));

            // 样式定义里正文是小四（24 半磅），Word 打开就是小四
            var normalStyle = ReadStyle(outputPath, "Normal");
            Assert.Equal("24", normalStyle.Element(W("rPr"))!.Element(W("sz"))!.Attribute(W("val"))!.Value);
        }
        finally
        {
            TryDelete(path);
            TryDelete(outputPath);
            try { Directory.Delete(backupDir, true); } catch (IOException) { }
        }
    }

    // ---------- 测试辅助 ----------

    private static readonly System.Reflection.Assembly CliAssembly = LoadCliAssembly();

    private static object NewStyleAnalyzeCommand(System.Reflection.Assembly cli, TextWriter output, TextWriter error)
    {
        var type = cli.GetType("ThesisEngine.Cli.StyleAnalyzeCommand")!;
        return System.Activator.CreateInstance(type, null, output, error)!;
    }

    private static object NewFullFormatCommand(System.Reflection.Assembly cli, TextWriter output, TextWriter error, string backupDir)
    {
        var type = cli.GetType("ThesisEngine.Cli.FullFormatCommand")!;
        var undo = new UndoManager(backupDir);
        return System.Activator.CreateInstance(type, new object?[] { null, null, undo, output, error })!;
    }

    private static XDocument ReadDocumentXml(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = archive.GetEntry("word/document.xml")!;
        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }

    /// <summary>取正文里第一个带文字的 run 的 rPr。</summary>
    private static XElement? FirstRunProperties(string docxPath) =>
        ReadDocumentXml(docxPath).Root!
            .Element(W("body"))!
            .Descendants(W("r"))
            .Where(r => r.Element(W("t")) != null)
            .Select(r => r.Element(W("rPr")))
            .FirstOrDefault();

    /// <summary>读文档里某个 styleId 对应的样式定义。</summary>
    private static XElement ReadStyle(string docxPath, string styleId)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = archive.GetEntry("word/styles.xml")!;
        using var reader = new StreamReader(entry.Open());
        var doc = XDocument.Load(reader);
        return doc.Root!.Elements(W("style"))
            .Single(s => s.Attribute(W("styleId"))?.Value == styleId);
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
