using System.Diagnostics;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DemoEndToEndTests
{
    private const string CleanSample = "tests/fixtures/synthetic/demo/clean_sample.docx";
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
    public async Task FullSyntheticThesisFindsOnlyItsThreePlannedBodyFaults()
    {
        var sample = Path.Combine(FindRepositoryRoot(), FullThesis);
        var before = File.ReadAllBytes(sample);

        var result = await RunCli("analyze", "--demo", FullThesis);

        Assert.Equal(0, result.ExitCode);
        Assert.Empty(result.Error);
        Assert.Contains("诊断数量：3", result.Output);
        Assert.Contains("已检查 47 个正文文本片段，涉及 41 个确认正文段落", result.Output);
        Assert.Contains("demo/east-asian-font", result.Output);
        Assert.Contains("demo/western-font", result.Output);
        Assert.Contains("demo/font-size", result.Output);
        Assert.Contains("表格单元格", result.Output);
        Assert.Contains("figure/table caption", result.Output);
        Assert.Contains("images", result.Output);
        Assert.Equal(before, File.ReadAllBytes(sample));
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
