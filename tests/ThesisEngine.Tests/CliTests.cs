using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class CliTests
{
    [Theory]
    [InlineData("corrupt", "")]
    [InlineData("missing-part", "Missing or duplicate OOXML part: [Content_Types].xml")]
    [InlineData("duplicate-part", "Missing or duplicate OOXML part: [Content_Types].xml")]
    [InlineData("invalid-root", "Invalid OOXML root: [Content_Types].xml")]
    [InlineData("missing-body", "Missing Word document body.")]
    [InlineData("malformed-xml", "")]
    public async Task InvalidPackagesReturnReadableErrorWithoutChangingInput(string scenario, string detail)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            CreatePackage(path, scenario);
            var before = File.ReadAllBytes(path);
            var result = await RunCli(path);
            Assert.Equal(1, result.ExitCode);
            Assert.StartsWith("Unable to load document:", result.Error);
            Assert.Contains(detail, result.Error);
            Assert.DoesNotContain("Unhandled exception", result.Error);
            Assert.Empty(result.Output);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public async Task ValidPackageLoadsWithPlaceholderWarning()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            CreatePackage(path, "valid");
            var before = File.ReadAllBytes(path);
            var result = await RunCli(path);
            Assert.Equal(0, result.ExitCode);
            Assert.Contains("Document loaded successfully", result.Output);
            Assert.Contains("不能据此判断合规", result.Output);
            Assert.Empty(result.Error);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public async Task MissingArgumentReturnsUsage()
    {
        var result = await RunCli();
        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Usage: ThesisEngine.Cli", result.Error);
        Assert.Empty(result.Output);
    }

    [Fact]
    public async Task MissingFileReturnsReadableError()
    {
        var result = await RunCli(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx"));
        Assert.Equal(1, result.ExitCode);
        Assert.StartsWith("Unable to load document:", result.Error);
        Assert.DoesNotContain("Unhandled exception", result.Error);
        Assert.Empty(result.Output);
    }

    [Fact]
    public async Task AnalyzeWithAnnotatedOutputWritesCommentsAndPreservesInput()
    {
        var input = GetDemoFixture("known_issues_sample.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".annotated.docx");
        var before = File.ReadAllBytes(input);
        try
        {
            var result = await RunCli("analyze", "--demo", input, "--annotated-output", output);

            Assert.Equal(0, result.ExitCode);
            Assert.Empty(result.Error);
            Assert.Contains("诊断数量：3", result.Output);
            Assert.Contains("规则版本：demo-0.1.0", result.Output);
            Assert.Contains("已检查范围：", result.Output);
            Assert.Contains("未覆盖范围：", result.Output);
            Assert.Contains($"标注副本：{output}", result.Output);
            Assert.Contains("原文件不变", result.Output);
            Assert.Equal(before, File.ReadAllBytes(input));
            Assert.True(File.Exists(output));

            using var archive = ZipFile.OpenRead(output);
            var comments = ReadXml(archive, "word/comments.xml");
            var document = ReadXml(archive, "word/document.xml");
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            Assert.Equal(3, comments.Root!.Elements(word + "comment").Count());
            Assert.Equal(3, document.Descendants(word + "highlight").Count());
            var commentText = string.Concat(comments.Descendants(word + "t").Select(element => element.Value));
            Assert.Contains("宋体", commentText);
            Assert.Contains("Times New Roman", commentText);
            Assert.Contains("12pt", commentText);
        }
        finally
        {
            File.Delete(output);
        }
    }

    [Fact]
    public async Task AnalyzeCleanSampleCreatesUnannotatedCopyAndReportsCoverage()
    {
        var input = GetDemoFixture("clean_sample.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".clean.docx");
        var before = File.ReadAllBytes(input);
        try
        {
            var result = await RunCli("analyze", "--demo", input, "--annotated-output", output);

            Assert.Equal(0, result.ExitCode);
            Assert.Empty(result.Error);
            Assert.Contains("诊断数量：0", result.Output);
            Assert.Contains("规则版本：demo-0.1.0", result.Output);
            Assert.Contains("已检查范围：", result.Output);
            Assert.Contains("未覆盖范围：", result.Output);
            Assert.Contains($"标注副本：{output}", result.Output);
            Assert.Equal(before, File.ReadAllBytes(input));
            Assert.True(File.Exists(output));
            using var archive = ZipFile.OpenRead(output);
            Assert.Null(archive.GetEntry("word/comments.xml"));
        }
        finally
        {
            File.Delete(output);
        }
    }

    [Fact]
    public async Task AnalyzeRejectsSameInputAndOutputPathWithoutChangingInput()
    {
        var input = GetDemoFixture("known_issues_sample.docx");
        var before = File.ReadAllBytes(input);

        var result = await RunCli("analyze", "--demo", input, "--annotated-output", input);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("输出路径不能与输入路径相同", result.Error);
        Assert.Empty(result.Output);
        Assert.Equal(before, File.ReadAllBytes(input));
    }

    [Fact]
    public async Task AnalyzeDoesNotOverwriteExistingOutput()
    {
        var input = GetDemoFixture("known_issues_sample.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".existing.docx");
        var sentinel = new byte[] { 1, 3, 3, 7 };
        File.WriteAllBytes(output, sentinel);
        try
        {
            var result = await RunCli("analyze", "--demo", input, "--annotated-output", output);

            Assert.Equal(1, result.ExitCode);
            Assert.Contains("目标已存在", result.Error);
            Assert.Empty(result.Output);
            Assert.Equal(sentinel, File.ReadAllBytes(output));
        }
        finally
        {
            File.Delete(output);
        }
    }

    [Fact]
    public async Task AnalyzeRejectsInvalidOutputExtensionAndMissingDirectory()
    {
        var input = GetDemoFixture("known_issues_sample.docx");
        var invalidExtension = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".txt");
        var missingDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "result.docx");

        var extensionResult = await RunCli("analyze", "--demo", input, "--annotated-output", invalidExtension);
        var directoryResult = await RunCli("analyze", "--demo", input, "--annotated-output", missingDirectory);

        Assert.Equal(1, extensionResult.ExitCode);
        Assert.Contains("标注副本必须使用 .docx", extensionResult.Error);
        Assert.Empty(extensionResult.Output);
        Assert.Equal(1, directoryResult.ExitCode);
        Assert.Contains("标注输出目录不存在", directoryResult.Error);
        Assert.Empty(directoryResult.Output);
        Assert.False(File.Exists(invalidExtension));
        Assert.False(File.Exists(missingDirectory));
    }

    [Fact]
    public async Task AnalyzeReportsInvalidArgumentsAndMissingDocumentsInChinese()
    {
        var malformed = await RunCli("analyze", "--demo", "document.docx", "--annotated-output");
        Assert.Equal(2, malformed.ExitCode);
        Assert.Contains("用法：", malformed.Error);
        Assert.Empty(malformed.Output);

        var missingPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var missing = await RunCli("analyze", "--demo", missingPath);
        Assert.Equal(1, missing.ExitCode);
        Assert.Contains("找不到输入文件", missing.Error);
        Assert.Empty(missing.Output);
    }

    private static async Task<(int ExitCode, string Output, string Error)> RunCli(params string[] args)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln")))
            root = root.Parent;
        Assert.NotNull(root);
        var testOutput = new DirectoryInfo(AppContext.BaseDirectory);
        var configuration = testOutput.Parent!.Name;
        var targetFramework = testOutput.Name;
        var cli = Path.Combine(root.FullName, "src", "ThesisEngine.Cli", "bin", configuration, targetFramework, "ThesisEngine.Cli.dll");
        Assert.True(File.Exists(cli), $"CLI was not built: {cli}");
        var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        start.ArgumentList.Add(cli);
        foreach (var arg in args) start.ArgumentList.Add(arg);
        using var process = Process.Start(start)!;
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            throw;
        }
        return (process.ExitCode, await output, await error);
    }

    private static string GetDemoFixture(string fileName)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln")))
            root = root.Parent;
        Assert.NotNull(root);
        return Path.Combine(root.FullName, "tests", "fixtures", "synthetic", "demo", fileName);
    }

    private static XDocument ReadXml(ZipArchive archive, string partName)
    {
        var entry = archive.GetEntry(partName);
        Assert.NotNull(entry);
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static void CreatePackage(string path, string scenario)
    {
        if (scenario == "corrupt")
        {
            File.WriteAllText(path, "not a ZIP archive");
            return;
        }
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        if (scenario == "missing-part") return;
        var types = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>";
        Add(archive, "[Content_Types].xml", scenario switch
        {
            "invalid-root" => "<Wrong/>",
            "malformed-xml" => "<Types",
            _ => types
        });
        if (scenario == "duplicate-part") Add(archive, "[Content_Types].xml", types);
        Add(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        Add(archive, "word/document.xml", scenario == "missing-body"
            ? "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"
            : "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>");
    }

    private static void Add(ZipArchive archive, string name, string xml)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(xml);
    }
}
