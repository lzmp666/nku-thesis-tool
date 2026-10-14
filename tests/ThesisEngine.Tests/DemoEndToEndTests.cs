using System.Diagnostics;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DemoEndToEndTests
{
    private const string CleanSample = "tests/fixtures/synthetic/demo/clean_sample.docx";
    private const string KnownIssuesSample = "tests/fixtures/synthetic/demo/known_issues_sample.docx";

    [Fact]
    public async Task CleanSyntheticSampleProducesNoFindingsWithoutChangingInput()
    {
        var sample = Path.Combine(FindRepositoryRoot(), CleanSample);
        var before = File.ReadAllBytes(sample);

        var result = await RunCli("analyze", "--demo", CleanSample);

        Assert.Equal(0, result.ExitCode);
        Assert.Empty(result.Error);
        Assert.Contains("=== Thesis format analysis ===", result.Output);
        Assert.Contains("Mode: DRAFT/DEMO", result.Output);
        Assert.Contains("Diagnostics: 0", result.Output);
        Assert.Contains("Findings: 0", result.Output);
        Assert.Contains("Rule version: demo-0.1.0", result.Output);
        Assert.Contains("Finding details:\n  (none)", result.Output);
        Assert.Contains("仅供演示/测试，不代表合规结论", result.Output);
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
        Assert.Contains("Mode: DRAFT/DEMO", result.Output);
        Assert.Contains("Diagnostics: 3", result.Output);
        Assert.Contains("Findings: 3", result.Output);
        Assert.Contains("Rule version: demo-0.1.0", result.Output);
        Assert.Contains("demo/east-asian-font", result.Output);
        Assert.Contains("word/document.xml#paragraph[0]/run[0]", result.Output);
        Assert.Contains("Expected: 宋体", result.Output);
        Assert.Contains("Actual: 黑体", result.Output);
        Assert.Contains("demo/western-font", result.Output);
        Assert.Contains("word/document.xml#paragraph[1]/run[0]", result.Output);
        Assert.Contains("Expected: Times New Roman", result.Output);
        Assert.Contains("Actual: Arial", result.Output);
        Assert.Contains("demo/font-size", result.Output);
        Assert.Contains("word/document.xml#paragraph[2]/run[0]", result.Output);
        Assert.Contains("Expected: 12pt", result.Output);
        Assert.Contains("Actual: 10.5pt", result.Output);
        Assert.Contains("仅供演示/测试，不代表合规结论", result.Output);
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
