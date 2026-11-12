using System.IO.Compression;
using System.Xml.Linq;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// Issue #108 落地的三条主干回归基线：
/// 1. <b>修复后复检 0 条</b>——full-format 跑过的文档再 analyze，表格不再误报「有完整内边框」，
///    也不残留被样式接管的手动字体/字号；
/// 2. <b>二次运行幂等</b>——对同一输入跑两次 full-format，比的是解包后的
///    <c>word/document.xml</c>，不是整个文件的哈希（zip 条目时间戳每次都不同，按字节比永远过不了）；
/// 3. <b>继承边框的表格能被识别</b>——没有显式 <c>tblBorders</c>、边框可能来自表格样式的表格，
///    检查器必须明确提示「需确认」，不能悄悄当成三线表放行。
/// </summary>
public sealed class RegressionBaselineTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ---------- 基线 1：修复后复检 0 条 ----------

    [Fact]
    public void FullFormatThenRecheckReportsZeroTableIssues()
    {
        var root = FindRepositoryRoot();
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_object_layout_errors.docx");
        var output = NewTempDocxPath();
        var backupDir = Path.Combine(Path.GetTempPath(), "nku-undo-regression-" + Guid.NewGuid().ToString("N"));
        try
        {
            var analyzer = new FigureTableAnalyzer();
            Assert.NotEmpty(analyzer.Analyze(input).TableIssues);

            var exitCode = RunFullFormat(input, output, backupDir);
            Assert.Equal(0, exitCode);

            var after = analyzer.Analyze(output);
            Assert.Equal(analyzer.Analyze(input).TablesCount, after.TablesCount);
            Assert.Empty(after.TableIssues);
            Assert.Empty(after.PageBreakIssues);
        }
        finally
        {
            TryDelete(output);
            try { Directory.Delete(backupDir, true); } catch (IOException) { }
        }
    }

    [Fact]
    public void FullFormatClearsFontAndSizeOverridesSoStyleCheckStaysClean()
    {
        var root = FindRepositoryRoot();
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_object_layout_errors.docx");
        var output = NewTempDocxPath();
        var backupDir = Path.Combine(Path.GetTempPath(), "nku-undo-regression-" + Guid.NewGuid().ToString("N"));
        try
        {
            Assert.Equal(0, RunFullFormat(input, output, backupDir));

            // full-format 之后不允许再残留「手动字体/手动字号」这类误报；
            // 段前段后、对齐保留在段落上属有意行为（标题靠它们撑版面），不算误报。
            var overrides = new StyleAnalyzer().Analyze(output).DirectFormatOverrides;
            Assert.DoesNotContain(overrides, item => item.Overrides.Contains("手动字体"));
            Assert.DoesNotContain(overrides, item => item.Overrides.Contains("手动字号"));
        }
        finally
        {
            TryDelete(output);
            try { Directory.Delete(backupDir, true); } catch (IOException) { }
        }
    }

    // ---------- 基线 2：二次运行幂等（比 document.xml，不比文件哈希） ----------

    [Fact]
    public void FullFormatSecondRunProducesIdenticalDocumentXml()
    {
        var root = FindRepositoryRoot();
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_object_layout_errors.docx");
        var first = NewTempDocxPath();
        var second = NewTempDocxPath();
        var backupDir = Path.Combine(Path.GetTempPath(), "nku-undo-regression-" + Guid.NewGuid().ToString("N"));
        try
        {
            Assert.Equal(0, RunFullFormat(input, first, backupDir));
            Assert.Equal(0, RunFullFormat(input, second, backupDir));

            // 两个输出 docx 的整体哈希几乎必然不同：zip 条目时间戳取当前时刻。
            // 幂等的正确比法是解包 document.xml、压掉空白后比内容（修复不引入任何内容差异）。
            Assert.NotEqual(first, second);
            Assert.Equal(ReadCanonicalDocumentXml(first), ReadCanonicalDocumentXml(second));
        }
        finally
        {
            TryDelete(first);
            TryDelete(second);
            try { Directory.Delete(backupDir, true); } catch (IOException) { }
        }
    }

    // ---------- 基线 3：继承边框的表格能被识别 ----------

    [Fact]
    public void TableInheritingBordersFromStyleIsFlaggedForConfirmation()
    {
        var path = CreateDocxWithStyledTable();
        try
        {
            var result = new FigureTableAnalyzer().Analyze(path);

            Assert.Equal(1, result.TablesCount);
            var issue = Assert.Single(result.TableIssues);
            // 不显式写 tblBorders 时边框可能继承自表格样式，检查器必须说出来，
            // 而不是当成三线表悄悄放行，也不能误报成「有完整内边框」。
            Assert.Contains("未显式设置三线表边框", issue);
            Assert.Contains("继承自表格样式", issue);
        }
        finally { File.Delete(path); }
    }

    // ---------- 测试辅助 ----------

    private static int RunFullFormat(string input, string output, string backupDir)
    {
        var cli = LoadCliAssembly();
        var type = cli.GetType("ThesisEngine.Cli.FullFormatCommand")!;
        var command = Activator.CreateInstance(type, new object?[] { null, null, new UndoManager(backupDir), null, null })!;
        var execute = type.GetMethod("Execute", new[] { typeof(string[]) })!;
        return (int)execute.Invoke(command, new object[] { new[] { input, output } })!;
    }

    private static string ReadCanonicalDocumentXml(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = archive.GetEntry("word/document.xml")!;
        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader).ToString(SaveOptions.DisableFormatting);
    }

    private static string CreateDocxWithStyledTable()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml",
            """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>""");
        Add(archive, "_rels/.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""");
        Add(archive, "word/document.xml",
            $"""
             <w:document xmlns:w="{W}"><w:body>
               <w:tbl>
                 <w:tblPr><w:tblStyle w:val="NKUThreeLine"/></w:tblPr>
                 <w:tr><w:tc><w:p><w:r><w:t>项目</w:t></w:r></w:p></w:tc>
                 <w:tc><w:p><w:r><w:t>数值</w:t></w:r></w:p></w:tc></w:tr>
               </w:tbl>
             </w:body></w:document>
             """);
        return path;
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(content);
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

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static string NewTempDocxPath() =>
        Path.Combine(Path.GetTempPath(), "regression-tests-" + Guid.NewGuid().ToString("N") + ".docx");

    private static void TryDelete(string path)
    {
        try { File.Delete(path); }
        catch (IOException) { }
    }
}
