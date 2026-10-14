using System.Reflection;
using Xunit;

namespace ThesisEngine.Tests;

[Collection("CLI analyze command")]
public sealed class CliAnalyzeCommandTests
{
    [Fact]
    public void CorrectArgumentsAreParsedAndPassedToTheInjectedService()
    {
        var assembly = LoadCliAssembly();
        var result = NewResult(assembly);
        ServiceProxy.Result = result;
        ServiceProxy.Exception = null;
        ServiceProxy.LastPath = null;

        var path = CreateInputFile();
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var command = NewCommand(assembly, output, error);
            var exitCode = InvokeCommand(command, path, output, error);

            Assert.Equal(0, exitCode);
            Assert.Equal(path, ServiceProxy.LastPath);
            Assert.NotNull(ServiceProxy.LastProfile);
            var profileType = ServiceProxy.LastProfile!.GetType();
            Assert.True((bool)profileType.GetProperty("IsDraft")!.GetValue(ServiceProxy.LastProfile)!);
            Assert.Equal("demo-0.1.0", profileType.GetProperty("RuleVersion")!.GetValue(ServiceProxy.LastProfile));
            Assert.Contains("诊断数量：0", output.ToString());
            Assert.Contains("草稿/演示", output.ToString());
            Assert.DoesNotContain("全文合规", output.ToString(), StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void FindingsDoNotChangeSuccessfulExitCode()
    {
        var assembly = LoadCliAssembly();
        var result = NewResult(assembly, NewDiagnostic(assembly, "error"));
        ServiceProxy.Result = result;
        ServiceProxy.Exception = null;
        var path = CreateInputFile();
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var exitCode = InvokeCommand(NewCommand(assembly, output, error), path, output, error);

            Assert.Equal(0, exitCode);
            Assert.Contains("诊断数量：1", output.ToString());
            Assert.Contains("错误级别的问题：", output.ToString());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void MissingInputReturnsExitCodeOne()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Result = NewResult(assembly);
        ServiceProxy.Exception = null;
        ServiceProxy.LastPath = null;
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = InvokeCommand(NewCommand(assembly, output, error), path, output, error);

        Assert.Equal(1, exitCode);
        Assert.Contains("找不到输入文件", error.ToString());
        Assert.Null(ServiceProxy.LastPath);
    }

    [Fact]
    public void ServiceFailureReturnsInputOrAnalysisErrorCodeOne()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Exception = new IOException("cannot read document");
        var path = CreateInputFile();
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var exitCode = InvokeCommand(NewCommand(assembly, output, error), path, output, error);

            Assert.Equal(1, exitCode);
            Assert.Contains("无法读取文档", error.ToString());
        }
        finally
        {
            ServiceProxy.Exception = null;
            File.Delete(path);
        }
    }

    [Fact]
    public void InvalidUsageReturnsExitCodeTwo()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Result = NewResult(assembly);
        ServiceProxy.Exception = null;
        var output = new StringWriter();
        var error = new StringWriter();
        var command = NewCommand(assembly, output, error);
        var execute = command.GetType().GetMethod("Execute", new[] { typeof(string[]), typeof(TextWriter), typeof(TextWriter) })!;

        var missingPathExitCode = (int)execute.Invoke(command, new object[] { new[] { "analyze", "--demo" }, output, error })!;
        var misspelledCommandExitCode = (int)execute.Invoke(command, new object[] { new[] { "anlyze", "--demo", "sample.docx" }, output, error })!;

        Assert.Equal(2, missingPathExitCode);
        Assert.Equal(2, misspelledCommandExitCode);
        Assert.Contains("用法：ThesisEngine.Cli analyze --demo <file.docx>", error.ToString());
    }

    [Fact]
    public void ReportMarksDraftDemoAndNeverClaimsFullCompliance()
    {
        var assembly = LoadCliAssembly();
        var result = NewResult(assembly, new[] { NewDiagnostic(assembly) }, "body formatting", "tables");
        var formatter = assembly.GetType("ThesisEngine.Cli.CliReportFormatter")!;
        var format = formatter.GetMethod("Format", new[] { result.GetType(), typeof(string) })!;

        var report = (string)format.Invoke(null, new object[] { result, "sample.docx" })!;

        Assert.Contains("=== 论文格式检查报告 ===", report);
        Assert.Contains("模式：草稿/演示", report);
        Assert.Contains("问题详情：", report);
        Assert.Contains("位置：paragraph[0]/run[0]", report);
        Assert.Contains("期望：12pt", report);
        Assert.Contains("实际：10.5pt", report);
        Assert.Contains("修复级别：手动", report);
        Assert.Contains("说明：字号与演示要求不一致。", report);
        Assert.Contains("未覆盖范围：", report);
        Assert.Contains("tables", report);
        Assert.DoesNotContain("全文合规", report, StringComparison.Ordinal);
        Assert.DoesNotContain("全文合规", report, StringComparison.Ordinal);
    }

    private static object NewCommand(Assembly assembly, TextWriter output, TextWriter error)
    {
        var commandType = assembly.GetType("ThesisEngine.Cli.AnalyzeCommand")!;
        var serviceType = CoreType(assembly, "IDemoAnalysisService");
        var proxy = typeof(DispatchProxy)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Single(method => method.Name == "Create" &&
                              method.IsGenericMethodDefinition &&
                              method.GetGenericArguments().Length == 2 &&
                              method.GetParameters().Length == 0)
            .MakeGenericMethod(serviceType, typeof(ServiceProxy))
            .Invoke(null, null)!;
        return Activator.CreateInstance(commandType, proxy, output, error)!;
    }

    private static int InvokeCommand(object command, string path, TextWriter output, TextWriter error)
    {
        var execute = command.GetType().GetMethod("Execute", new[] { typeof(string[]), typeof(TextWriter), typeof(TextWriter) })!;
        return (int)execute.Invoke(command, new object[] { new[] { "analyze", "--demo", path }, output, error })!;
    }

    private static object NewResult(Assembly assembly, params object[] diagnostics)
    {
        return NewResult(assembly, diagnostics, "fonts and sizes", "body", "tables");
    }

    private static object NewResult(
        Assembly assembly,
        object[] diagnostics,
        string coverage,
        params string[] unsupported)
    {
        var resultType = CoreType(assembly, "AnalysisResult");
        var diagnosticType = CoreType(assembly, "Diagnostic");
        var diagnosticArray = Array.CreateInstance(diagnosticType, diagnostics.Length);
        for (var index = 0; index < diagnostics.Length; index++)
            diagnosticArray.SetValue(diagnostics[index], index);
        return Activator.CreateInstance(
            resultType,
            diagnosticArray,
            "demo-0.1.0",
            coverage,
            unsupported)!;
    }

    private static object NewDiagnostic(Assembly assembly, string severity = "warning")
    {
        var diagnosticType = CoreType(assembly, "Diagnostic");
        return Activator.CreateInstance(
            diagnosticType,
            "demo/font-size",
            severity,
            "paragraph[0]/run[0]",
            "12pt",
            "10.5pt",
            "manual",
            "字号与演示要求不一致。")!;
    }

    private static string CreateInputFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, Array.Empty<byte>());
        return path;
    }

    private static Assembly LoadCliAssembly()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln")))
            root = root.Parent;
        Assert.NotNull(root);

        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
        var path = Path.Combine(root!.FullName, "src", "ThesisEngine.Cli", "bin", configuration, "net10.0", "ThesisEngine.Cli.dll");
        Assert.True(File.Exists(path), $"CLI was not built: {path}");
        return Assembly.LoadFrom(path);
    }

    private static Type CoreType(Assembly cliAssembly, string typeName) =>
        cliAssembly.GetType($"ThesisEngine.Core.{typeName}") ??
        Assembly.Load("ThesisEngine.Core").GetType($"ThesisEngine.Core.{typeName}")!;

    public class ServiceProxy : DispatchProxy
    {
        public static object? Result { get; set; }
        public static Exception? Exception { get; set; }
        public static string? LastPath { get; set; }
        public static object? LastProfile { get; set; }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            LastPath = args?[0] as string;
            LastProfile = args?[1];
            if (Exception is not null)
                throw Exception;
            return Result;
        }
    }
}

[CollectionDefinition("CLI analyze command", DisableParallelization = true)]
public sealed class CliAnalyzeCommandCollection { }
