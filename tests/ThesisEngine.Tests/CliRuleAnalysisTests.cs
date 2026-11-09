using System.Reflection;
using System.Text.Json;
using ThesisEngine.Core.Analysis;
using Xunit;

namespace ThesisEngine.Tests;

[Collection("CLI analyze command")]
public sealed class CliRuleAnalysisTests
{
    [Fact]
    public void AnalyzeWithoutDemoUsesTheYamlRuleEngine()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var command = NewCommand(output, error);
        var exitCode = Invoke(command, new[] { "analyze", Fixture("analysis_compliant.docx") }, output, error);

        Assert.Equal(0, exitCode);
        Assert.Contains("=== 论文格式检查报告（规则分析）===", output.ToString());
        Assert.Contains("R11/font.cn", output.ToString());
        Assert.Contains("draft", output.ToString());
        Assert.DoesNotContain("占位模式", output.ToString(), StringComparison.Ordinal);
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public void AnalyzeJsonProducesMachineReadableOutput()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var command = NewCommand(output, error);
        var exitCode = Invoke(command, new[] { "analyze", Fixture("analysis_noncompliant.docx"), "--json" }, output, error);

        Assert.Equal(0, exitCode);
        using var document = JsonDocument.Parse(output.ToString());
        Assert.Equal("thesis-analysis-1", document.RootElement.GetProperty("schemaVersion").GetString());
        Assert.Equal(8, document.RootElement.GetProperty("findings").GetArrayLength());
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public void AnalyzeCanWriteAnAnnotatedCopy()
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"cli-annotated-{Guid.NewGuid():N}.docx");
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var command = NewCommand(output, error);
            var exitCode = Invoke(command, new[]
            {
                "analyze", Fixture("analysis_noncompliant.docx"), "--annotated-output", outputPath
            }, output, error);

            Assert.Equal(0, exitCode);
            Assert.True(File.Exists(outputPath));
            Assert.Contains("标注副本", output.ToString());
            Assert.Equal(string.Empty, error.ToString());
        }
        finally
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void AnalyzeAcceptsExplicitRuleFileAndProfile()
    {
        var ruleFile = Path.Combine(RepositoryRoot(), "rules", "nku-graduate", "2026-body-v1.yaml");
        var output = new StringWriter();
        var error = new StringWriter();
        var command = NewCommand(output, error);
        var exitCode = Invoke(command, new[]
        {
            "analyze", Fixture("analysis_noncompliant.docx"), "--rule-file", ruleFile
        }, output, error);

        Assert.Equal(0, exitCode);
        Assert.Contains("nku-graduate-2026-body", output.ToString());
        Assert.Contains("NKU-REQ-042", output.ToString());
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public void ConflictingOptionsReturnUsageError()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var command = NewCommand(output, error);
        var exitCode = Invoke(command, new[]
        {
            "analyze", Fixture("analysis_compliant.docx"), "--demo", "--profile", "nku-graduate"
        }, output, error);

        Assert.Equal(2, exitCode);
        Assert.Contains("用法：ThesisEngine.Cli analyze --demo <file.docx>", error.ToString());
    }

    private static object NewCommand(TextWriter output, TextWriter error)
    {
        var assembly = LoadCliAssembly();
        var commandType = assembly.GetType("ThesisEngine.Cli.AnalyzeCommand")!;
        return Activator.CreateInstance(commandType, new RuleAnalysisService(), output, error)!;
    }

    private static int Invoke(object command, string[] args, TextWriter output, TextWriter error)
    {
        var execute = command.GetType().GetMethod("Execute", new[] { typeof(string[]), typeof(TextWriter), typeof(TextWriter) })!;
        return (int)execute.Invoke(command, new object[] { args, output, error })!;
    }

    private static Assembly LoadCliAssembly()
    {
        var root = RepositoryRoot();
        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
        var path = Path.Combine(root, "src", "ThesisEngine.Cli", "bin", configuration, "net10.0", "ThesisEngine.Cli.dll");
        Assert.True(File.Exists(path), $"CLI was not built: {path}");
        return Assembly.LoadFrom(path);
    }

    private static string Fixture(string name) =>
        Path.Combine(RepositoryRoot(), "tests", "fixtures", "synthetic", name);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }
}
