using System.Diagnostics;
using System.Text;
using ThesisEngine.Core;
using Xunit;

/// <summary>
/// 统一合规报告：验证九类检查聚合、三类结果分列、失败隔离、JSON 输出。
/// 样本为 tests/fixtures/synthetic 下既有的零依赖合成 docx，不新增样本。
/// </summary>
public sealed class ComplianceReportTests
{
    private static string Root()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static string Fixture(params string[] parts) =>
        Path.Combine(new[] { Root(), "tests", "fixtures", "synthetic" }.Concat(parts).ToArray());

    private static ComplianceReport Report(string relativePath) =>
        new ComplianceReportService().Analyze(Fixture(relativePath.Split('/')));

    [Fact]
    public void AllTwelveAnalyzersRunAndReportSeparatesThreeBuckets()
    {
        var report = Report("abstract/abstract_compliant.docx");

        Assert.Equal(12, report.Sections.Count);
        Assert.True(report.AllSucceeded, "合规样本下十二个分析器都应成功执行");
        Assert.Empty(report.Conventions);
        Assert.True(report.TotalNotChecked > 0, "未判定项必须单列，不能靠空列表假装查全了");
        Assert.True(report.TotalIssues >= 0);
    }

    [Fact]
    public void ErrorSampleAggregatesIssuesFromMultipleAnalyzers()
    {
        var report = Report("full-thesis/synthetic_object_layout_errors.docx");

        Assert.True(report.AllSucceeded);
        var analyzers = report.Issues.Select(i => i.Analyzer).Distinct().ToList();
        Assert.True(analyzers.Count >= 3, $"错例应命中多类分析器，实际：{string.Join(",", analyzers)}");
        Assert.Contains("AbstractFormatAnalyzer", analyzers);
        Assert.Contains("BackMatterFormatAnalyzer", analyzers);
        Assert.True(report.TotalIssues > 0);
    }

    [Fact]
    public void FootnoteUndeterminedItemsAreNotCountedAsIssues()
    {
        var report = Report("abstract/abstract_compliant.docx");

        Assert.DoesNotContain(report.Issues, i => i.Code.EndsWith("Undetermined", StringComparison.Ordinal));
        Assert.Contains(report.NotChecked, x => x.Contains("序号形态", StringComparison.Ordinal));
    }

    [Fact]
    public void ConventionsAreExcludedFromSpecIssueCount()
    {
        var report = Report("full-thesis/synthetic_object_layout_errors.docx");

        Assert.Equal(report.Issues.Count, report.TotalIssues);
        Assert.Equal(report.Conventions.Count, report.TotalConventions);
        Assert.DoesNotContain(report.Issues, i => i.IsConvention);
        Assert.All(report.Conventions, i => Assert.True(i.IsConvention));
    }

    [Fact]
    public async Task CommandOutputsValidJsonWithPascalCaseFields()
    {
        var result = await RunCli("report", Fixture("abstract", "abstract_compliant.docx"), "--json");

        Assert.Equal(0, result.ExitCode);
        var json = result.Output;
        Assert.Contains("\"Sections\"", json, StringComparison.Ordinal);
        Assert.Contains("\"Issues\"", json, StringComparison.Ordinal);
        Assert.Contains("\"NotChecked\"", json, StringComparison.Ordinal);
        using var document = System.Text.Json.JsonDocument.Parse(json);
        Assert.True(document.RootElement.GetProperty("AllSucceeded").GetBoolean());
    }

    [Fact]
    public void BrokenDocxIsReportedWithoutCrashingWholeReport()
    {
        var broken = Path.Combine(Path.GetTempPath(), $"compliance-broken-{Guid.NewGuid():N}.docx");
        File.WriteAllText(broken, "not a docx");
        try
        {
            var report = new ComplianceReportService().Analyze(broken);

            Assert.Equal(12, report.Sections.Count);
            Assert.True(report.FailedSections > 0, "损坏文件至少要让部分检查器失败");
        }
        finally
        {
            File.Delete(broken);
        }
    }

    [Fact]
    public async Task CommandReturnsUsageCodeForMissingArgument()
    {
        var result = await RunCli("report");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("用法", result.Output, StringComparison.Ordinal);
    }

    private static async Task<(int ExitCode, string Output, string Error)> RunCli(params string[] args)
    {
        var testOutput = new DirectoryInfo(AppContext.BaseDirectory);
        var configuration = testOutput.Parent!.Name;
        var targetFramework = testOutput.Name;
        var cli = Path.Combine(Root(), "src", "ThesisEngine.Cli", "bin", configuration, targetFramework, "ThesisEngine.Cli.dll");
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
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            throw;
        }
        return (process.ExitCode, await output, await error);
    }
}
