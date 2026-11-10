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
        var result = NewResult(assembly, new[] { NewDiagnostic(assembly) }, "正文格式", "tables");
        var formatter = assembly.GetType("ThesisEngine.Cli.CliReportFormatter")!;
        var format = formatter.GetMethod("Format", new[] { result.GetType(), typeof(string) })!;

        var report = (string)format.Invoke(null, new object[] { result, "sample.docx" })!;

        Assert.Contains("=== 论文格式检查报告 ===", report);
        Assert.Contains("模式：草稿/演示", report);
        Assert.Contains("警告 demo/font-size", report);
        Assert.Contains("问题详情：", report);
        Assert.Contains("位置：paragraph[0]/run[0]", report);
        Assert.Contains("期望：12pt", report);
        Assert.Contains("实际：10.5pt", report);
        Assert.Contains("修复级别：手动", report);
        Assert.Contains("说明：字号与演示要求不一致。", report);
        Assert.Contains("未覆盖范围：", report);
        Assert.Contains("表格", report);
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
        return NewResult(assembly, diagnostics, "字体和字号", "正文", "tables");
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

    // ---- #112：analyze --scope 的三种口径 ----

    [Fact]
    public void DocumentScopeSkipsParagraphAnalysis()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Result = NewResult(assembly);
        ServiceProxy.Exception = null;
        ServiceProxy.LastPath = null;
        var path = CreateInputFile();
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var exitCode = InvokeWithArgs(NewCommand(assembly, output, error),
                new[] { "analyze", "--demo", path, "--scope", "document" }, output, error);

            Assert.Null(ServiceProxy.LastPath);
            Assert.DoesNotContain("诊断数量", output.ToString());
            // 空文件不是有效 docx，文档级会自己报错，不该被段落级的结果盖过去。
            Assert.Equal(1, exitCode);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DocumentScopeRendersRealDocumentReport()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Result = NewResult(assembly);
        ServiceProxy.Exception = null;
        ServiceProxy.LastPath = null;

        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = InvokeWithArgs(NewCommand(assembly, output, error),
            new[] { "analyze", "--demo", SamplePath("document_baseline_correct.docx"), "--scope", "document" },
            output, error);

        Assert.Equal(0, exitCode);
        Assert.Null(ServiceProxy.LastPath);
        Assert.Contains("=== 文档级版式检查报告 ===", output.ToString());
        Assert.Contains("已检查的文档级项目均符合。", output.ToString());
        Assert.DoesNotContain("诊断数量", output.ToString());
    }

    [Fact]
    public void ParagraphScopeKeepsOnlyTheOriginalReport()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Result = NewResult(assembly);
        ServiceProxy.Exception = null;
        var path = CreateInputFile();
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var exitCode = InvokeWithArgs(NewCommand(assembly, output, error),
                new[] { "analyze", "--demo", path, "--scope", "paragraph" }, output, error);

            Assert.Equal(0, exitCode);
            Assert.Contains("诊断数量：0", output.ToString());
            Assert.DoesNotContain("文档级版式检查报告", output.ToString());
        }
        finally { File.Delete(path); }
    }

    /// <summary>默认 all 时文档级是附加报告，它失败不该把段落级的成功改成失败。</summary>
    [Fact]
    public void DefaultScopeAppendsDocumentSectionWithoutChangingExitCode()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Result = NewResult(assembly);
        ServiceProxy.Exception = null;
        var path = CreateInputFile();
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var exitCode = InvokeWithArgs(NewCommand(assembly, output, error),
                new[] { "analyze", "--demo", path }, output, error);

            Assert.Equal(0, exitCode);
            Assert.Contains("诊断数量：0", output.ToString());
            Assert.Contains("以下为文档级版式检查", output.ToString());
            Assert.Contains("文档级检查未完成", output.ToString());
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void InvalidScopeReturnsUsage()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Result = NewResult(assembly);
        var path = CreateInputFile();
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var exitCode = InvokeWithArgs(NewCommand(assembly, output, error),
                new[] { "analyze", "--demo", path, "--scope", "unknown" }, output, error);

            Assert.Equal(2, exitCode);
            Assert.Contains("--scope all|paragraph|document", error.ToString());
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DocumentScopeRejectsAnnotatedOutput()
    {
        var assembly = LoadCliAssembly();
        ServiceProxy.Result = NewResult(assembly);
        var path = CreateInputFile();
        try
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var exitCode = InvokeWithArgs(NewCommand(assembly, output, error),
                new[] { "analyze", "--demo", path, "--scope", "document",
                    "--annotated-output", Path.Combine(Path.GetTempPath(), "unused.docx") },
                output, error);

            Assert.Equal(2, exitCode);
            Assert.Contains("没有段落级诊断可标注", error.ToString());
        }
        finally { File.Delete(path); }
    }

    private static int InvokeWithArgs(object command, string[] args, TextWriter output, TextWriter error)
    {
        var execute = command.GetType().GetMethod("Execute",
            new[] { typeof(string[]), typeof(TextWriter), typeof(TextWriter) })!;
        return (int)execute.Invoke(command, new object[] { args, output, error })!;
    }

    private static string SamplePath(string sampleName)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln")))
            root = root.Parent;
        Assert.NotNull(root);
        var path = Path.Combine(root!.FullName, "tests", "fixtures", "synthetic", "document", sampleName);
        Assert.True(File.Exists(path), $"样本不存在：{path}");
        return path;
    }

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
