using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DemoEndToEndTests
{
    private const string CleanSample = "tests/fixtures/synthetic/demo/clean_sample.docx";
    private const string CleanCheckedSample = "tests/fixtures/synthetic/demo/clean_checked_sample.docx";
    private const string KnownIssuesSample = "tests/fixtures/synthetic/demo/known_issues_sample.docx";
    private const string FullThesis = "tests/fixtures/synthetic/full-thesis/synthetic_complete_thesis.docx";

    [Fact]
    public async Task CleanSyntheticSampleProducesNoFindingsWithoutChangingInput()
    {
        var sample = Path.Combine(FindRepositoryRoot(), CleanSample);
        var before = File.ReadAllBytes(sample);

        var result = await RunCli("analyze", "--demo", CleanSample);

        Assert.Equal(0, result.ExitCode);
        Assert.Empty(result.Error);
        Assert.Contains("=== 论文格式检查报告 ===", result.Output);
        Assert.Contains("模式：草稿/演示", result.Output);
        Assert.Contains("诊断数量：0", result.Output);
        Assert.Contains("发现问题：0", result.Output);
        Assert.Contains("规则版本：demo-0.1.0", result.Output);
        Assert.Contains("问题详情：", result.Output);
        Assert.Contains("（无）", result.Output);
        Assert.Contains("不代表规范合规结论", result.Output);
        Assert.Equal(before, File.ReadAllBytes(sample));
    }

    [Fact]
    public async Task KnownIssuesSampleProducesExactlyThreeExpectedFindingsWithoutChangingInput()
    {
        var sample = Path.Combine(FindRepositoryRoot(), KnownIssuesSample);
        var before = File.ReadAllBytes(sample);

        var result = await RunCli("analyze", "--demo", KnownIssuesSample);

        Assert.Equal(0, result.ExitCode);
        Assert.Empty(result.Error);
        Assert.Contains("模式：草稿/演示", result.Output);
        Assert.Contains("诊断数量：3", result.Output);
        Assert.Contains("发现问题：3", result.Output);
        Assert.Contains("规则版本：demo-0.1.0", result.Output);
        Assert.Contains("demo/east-asian-font", result.Output);
        Assert.Contains("word/document.xml#paragraph[1]/run[0]", result.Output);
        Assert.Contains("期望：宋体", result.Output);
        Assert.Contains("实际：黑体", result.Output);
        Assert.Contains("demo/western-font", result.Output);
        Assert.Contains("word/document.xml#paragraph[2]/run[0]", result.Output);
        Assert.Contains("期望：Times New Roman", result.Output);
        Assert.Contains("实际：Arial", result.Output);
        Assert.Contains("demo/font-size", result.Output);
        Assert.Contains("word/document.xml#paragraph[3]/run[0]", result.Output);
        Assert.Contains("期望：12pt", result.Output);
        Assert.Contains("实际：10.5pt", result.Output);
        Assert.Contains("不代表规范合规结论", result.Output);
        Assert.Equal(before, File.ReadAllBytes(sample));
    }

    [Fact]
    public async Task FullSyntheticThesisFindsThreeBodyFaultsAndOneGroupedCaptionStyleSet()
    {
        var sample = Path.Combine(FindRepositoryRoot(), FullThesis);
        var before = File.ReadAllBytes(sample);

        var result = await RunCli("analyze", "--demo", FullThesis);

        Assert.Equal(0, result.ExitCode);
        Assert.Empty(result.Error);
        Assert.Contains("诊断数量：9", result.Output);
        Assert.Contains("已检查 49 个正文文本片段，涉及 43 个确认正文段落", result.Output);
        Assert.Contains("识别 6 条图/表题注", result.Output);
        Assert.Contains("demo/east-asian-font", result.Output);
        Assert.Contains("demo/western-font", result.Output);
        Assert.Contains("demo/font-size", result.Output);
        Assert.Contains("demo/caption-east-asian-font", result.Output);
        Assert.Contains("demo/caption-font-size", result.Output);
        Assert.Contains("demo/caption-alignment", result.Output);
        Assert.Contains("demo/caption-line-spacing", result.Output);
        Assert.Contains("demo/caption-space-before", result.Output);
        Assert.Contains("demo/caption-space-after", result.Output);
        Assert.Contains("tables (cell formatting rules are not configured)", result.Output);
        Assert.Contains("images", result.Output);
        Assert.Equal(before, File.ReadAllBytes(sample));
    }

    // clean_sample.docx 的「0 条」来自「已检查 0 个正文文本片段」，是反面教材；
    // clean_checked_sample.docx 用来证明检查器确实查过，而且查下来没有问题。
    // 下面几条把这个区别钉死，防止样本退化回「压根没查」却仍被称为正确样例。
    [Fact]
    public async Task CleanCheckedSampleReportsZeroFindingsWithRealParagraphCoverage()
    {
        var sample = Path.Combine(FindRepositoryRoot(), CleanCheckedSample);
        var before = File.ReadAllBytes(sample);

        var result = await RunCli("analyze", "--demo", CleanCheckedSample, "--scope", "paragraph");

        Assert.Equal(0, result.ExitCode);
        Assert.Empty(result.Error);
        Assert.Contains("诊断数量：0", result.Output);
        Assert.Contains("发现问题：0", result.Output);
        Assert.Contains("已检查 3 个正文文本片段，涉及 3 个确认正文段落", result.Output);
        Assert.DoesNotContain("已检查 0 个正文文本片段", result.Output);
        Assert.DoesNotContain("正文范围未确认", result.Output);
        Assert.Equal(before, File.ReadAllBytes(sample));
    }

    [Fact]
    public async Task CleanCheckedSamplePassesAllDocumentLevelChecks()
    {
        var result = await RunCli("analyze", "--demo", CleanCheckedSample, "--scope", "document");

        Assert.Equal(0, result.ExitCode);
        Assert.Empty(result.Error);
        Assert.Contains("完整判定 2 / 2 个文字段落", result.Output);
        Assert.Contains("已检查的文档级项目均符合。", result.Output);
        Assert.DoesNotContain("⚠️ 不符合规范", result.Output);
    }

    [Fact]
    public async Task CleanCheckedSampleDefaultScopeCoversBothLevels()
    {
        var result = await RunCli("analyze", "--demo", CleanCheckedSample);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("=== 论文格式检查报告 ===", result.Output);
        Assert.Contains("=== 文档级版式检查报告 ===", result.Output);
        Assert.Contains("已检查 3 个正文文本片段", result.Output);
        Assert.Contains("已检查的文档级项目均符合。", result.Output);
        Assert.DoesNotContain("⚠️ 不符合规范", result.Output);
    }

    [Fact]
    public void CleanCheckedSamplePackageKeepsBothLevelPrerequisites()
    {
        var path = Path.Combine(FindRepositoryRoot(), CleanCheckedSample);
        using var archive = ZipFile.OpenRead(path);
        var names = archive.Entries.Select(entry => entry.FullName).ToArray();

        Assert.Contains("word/styles.xml", names);
        Assert.Contains("word/header1.xml", names);
        Assert.Contains("word/footer1.xml", names);

        var document = ReadPart(archive, "word/document.xml");
        Assert.Contains("Heading1", document);
        Assert.Contains("w:top=\"2155\"", document);
        Assert.Contains("w:header=\"1701\"", document);
        Assert.Contains("PAGE", ReadPart(archive, "word/footer1.xml"));
        Assert.Contains("第一章 绪论", ReadPart(archive, "word/header1.xml"));
    }

    [Fact]
    public async Task LegacyCleanSampleStillShowsZeroCheckedFragments()
    {
        var result = await RunCli("analyze", "--demo", CleanSample, "--scope", "paragraph");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("诊断数量：0", result.Output);
        Assert.Contains("已检查 0 个正文文本片段", result.Output);
        Assert.Contains("正文范围未确认", result.Output);
    }

    private static string ReadPart(ZipArchive archive, string partName)
    {
        var entry = archive.GetEntry(partName);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static async Task<(int ExitCode, string Output, string Error)> RunCli(params string[] arguments)
    {
        var root = FindRepositoryRoot();
        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent?.Name ?? "Release";
        var cliPath = Path.Combine(root, "src", "ThesisEngine.Cli", "bin", configuration, "net10.0", "ThesisEngine.Cli.dll");
        Assert.True(File.Exists(cliPath), $"CLI was not built: {cliPath}");

        var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet")
        {
            WorkingDirectory = root,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        start.ArgumentList.Add(cliPath);
        foreach (var argument in arguments)
            start.ArgumentList.Add(argument);

        using var process = Process.Start(start)!;
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try
        {
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            throw;
        }

        return (process.ExitCode, await output, await error);
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;

        return directory?.FullName ?? throw new DirectoryNotFoundException("Could not find NkuThesisTool.sln.");
    }
}
