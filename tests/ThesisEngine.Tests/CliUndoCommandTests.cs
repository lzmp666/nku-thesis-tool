using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// Issue #95：批量修改的撤销功能 —— 命令层集成测试。
/// 流程：先快照 → 再修改 → 撤销时回滚快照。
/// Cli 程序集未作为引用注入，按既有测试约定通过反射加载。
/// </summary>
public sealed class CliUndoCommandTests
{
    [Fact]
    public void UndoCommandWithoutArgumentsReturnsUsage()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var command = NewUndoCommand(new UndoManager(NewTempDir()), output, error);

        var exitCode = Execute(command, Array.Empty<string>());

        Assert.Equal(2, exitCode);
        Assert.Contains("用法", error.ToString());
    }

    [Fact]
    public void UndoCommandWithMissingFileReturnsReadableError()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var command = NewUndoCommand(new UndoManager(NewTempDir()), output, error);

        var exitCode = Execute(command, new[] { Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx") });

        Assert.Equal(1, exitCode);
        Assert.Contains("找不到文件", error.ToString());
    }

    [Fact]
    public void UndoCommandWithoutBackupReturnsReadableError()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var path = Path.Combine(dir, "thesis.docx");
            File.WriteAllText(path, "original");
            var output = new StringWriter();
            var error = new StringWriter();
            var command = NewUndoCommand(new UndoManager(backupDir), output, error);

            var exitCode = Execute(command, new[] { path });

            Assert.Equal(1, exitCode);
            Assert.Contains("撤销失败", error.ToString());
            Assert.Contains("没有找到备份", error.ToString());
            Assert.Equal("original", File.ReadAllText(path));
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    [Fact]
    public void UndoListPrintsRecordedChanges()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var undo = new UndoManager(backupDir);
            var path = Path.Combine(dir, "thesis.docx");
            File.WriteAllText(path, "original");
            undo.BackupBeforeModify(path, changeDescription: "套用南开格式：正文宋体小四");

            var output = new StringWriter();
            var error = new StringWriter();
            var command = NewUndoCommand(undo, output, error);

            var exitCode = Execute(command, new[] { "--list", path });

            Assert.Equal(0, exitCode);
            Assert.Contains("套用南开格式：正文宋体小四", output.ToString());
            Assert.Contains(Path.GetFileName(path), output.ToString());
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    [Fact]
    public void ApplyNkuFormatCreatesSnapshotAndUndoRestoresOriginalDocument()
    {
        var dir = NewTempDir();
        var backupDir = NewTempDir();
        try
        {
            var input = Path.Combine(dir, "thesis.docx");
            var output = Path.Combine(dir, "thesis_nku.docx");
            CreateMinimalDocx(input);
            var before = File.ReadAllBytes(input);

            var undo = new UndoManager(backupDir);
            var outputWriter = new StringWriter();
            var error = new StringWriter();
            var command = NewApplyNkuFormatCommand(undo, outputWriter, error);

            var exitCode = Execute(command, new[] { input, output });

            Assert.Equal(0, exitCode);
            Assert.True(File.Exists(output));

            // 批量修改确实改变了文档，且备份清单记录了本次改动
            var modified = File.ReadAllBytes(output);
            Assert.NotEqual(before, modified);
            var record = Assert.Single(undo.ListBackups());
            Assert.Equal(Path.GetFullPath(output), record.OutputPath);
            Assert.Contains("南开格式", record.ChangeDescription);

            // 一键撤销：输出文件恢复到修改前的状态
            var undoOutput = new StringWriter();
            var undoError = new StringWriter();
            var undoCommand = NewUndoCommand(undo, undoOutput, undoError);
            var undoExitCode = Execute(undoCommand, new[] { output });

            Assert.Equal(0, undoExitCode);
            Assert.Equal(before, File.ReadAllBytes(output));
            Assert.Contains("已撤销最近一次批量修改", undoOutput.ToString());
        }
        finally
        {
            TryDelete(dir);
            TryDelete(backupDir);
        }
    }

    private static object NewUndoCommand(UndoManager undo, TextWriter output, TextWriter error)
    {
        var type = CliAssembly.GetType("ThesisEngine.Cli.UndoCommand")!;
        return Activator.CreateInstance(type, undo, output, error)!;
    }

    private static object NewApplyNkuFormatCommand(UndoManager undo, TextWriter output, TextWriter error)
    {
        var type = CliAssembly.GetType("ThesisEngine.Cli.ApplyNkuFormatCommand")!;
        return Activator.CreateInstance(type, null, undo, output, error)!;
    }

    private static int Execute(object command, string[] args)
    {
        var execute = command.GetType().GetMethod("Execute", new[] { typeof(string[]) })!;
        return (int)execute.Invoke(command, new object[] { args })!;
    }

    private static readonly Assembly CliAssembly = LoadCliAssembly();

    private static Assembly LoadCliAssembly()
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
        return Assembly.LoadFrom(dll);
    }

    private static string NewTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "undo-cli-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void TryDelete(string dir)
    {
        try { Directory.Delete(dir, recursive: true); }
        catch (IOException) { }
    }

    /// <summary>构造一个带 Normal / Heading1 / Heading2 样式的最小合法 docx。</summary>
    private static void CreateMinimalDocx(string path)
    {
        var styles = new XDocument(new XElement(
            W("styles"),
            new XElement(W("style"),
                new XAttribute(W("type"), "paragraph"),
                new XAttribute(W("default"), "1"),
                new XAttribute(W("styleId"), "Normal"),
                new XElement(W("name"), new XAttribute(W("val"), "Normal"))),
            new XElement(W("style"),
                new XAttribute(W("type"), "paragraph"),
                new XAttribute(W("styleId"), "Heading1"),
                new XElement(W("name"), new XAttribute(W("val"), "heading 1")),
                new XElement(W("rPr"), new XElement(W("sz"), new XAttribute(W("val"), "32")))),
            new XElement(W("style"),
                new XAttribute(W("type"), "paragraph"),
                new XAttribute(W("styleId"), "Heading2"),
                new XElement(W("name"), new XAttribute(W("val"), "heading 2")),
                new XElement(W("rPr"), new XElement(W("sz"), new XAttribute(W("val"), "28"))))));

        var document = new XDocument(new XElement(
            W("document"),
            new XElement(W("body"),
                new XElement(W("p"),
                    new XElement(W("pPr"), new XElement(W("pStyle"), new XAttribute(W("val"), "Heading1"))),
                    new XElement(W("r"), new XElement(W("t"), "第一章 绪论"))),
                new XElement(W("p"),
                    new XElement(W("r"), new XElement(W("t"), "这是正文内容。"))))));

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
        WriteEntry(zip, "word/styles.xml", styles.ToString(SaveOptions.DisableFormatting));
    }

    private static void WriteEntry(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static XName W(string localName) =>
        (XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main" + localName;
}
