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
            Assert.Equal("draft/demo", profileType.GetProperty("DraftMarker")!.GetValue(ServiceProxy.LastProfile));
            Assert.Contains("Diagnostics: 0", output.ToString());
            Assert.Contains("DRAFT/DEMO", output.ToString());
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
            Assert.Contains("Diagnostics: 1", output.ToString());
            Assert.Contains("Error-severity findings:", output.ToString());
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
        Assert.Contains("file not found", error.ToString());
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
            Assert.Contains("cannot read document", error.ToString());
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
        Assert.Contains("Usage: ThesisEngine.Cli analyze --demo <file.docx>", error.ToString());
    }

    [Fact]
    public void ReportMarksDraftDemoAndNeverClaimsFullCompliance()
    {
        var assembly = LoadCliAssembly();
        var result = NewResult(assembly, new[] { NewDiagnostic(assembly) }, "body formatting", "tables");
        var formatter = assembly.GetType("ThesisEngine.Cli.CliReportFormatter")!;
        var format = formatter.GetMethod("Format", new[] { result.GetType(), typeof(string) })!;

        var report = (string)format.Invoke(null, new object[] { result, "sample.docx" })!;

        Assert.Contains("DRAFT/DEMO", report);
        Assert.Contains("Unsupported scopes", report);
        Assert.Contains("tables", report);
        Assert.DoesNotContain("full compliance", report, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("全文合规", report, StringComparison.Ordinal);
    }

    private static object NewCommand(Assembly assembly, TextWriter output, TextWriter error)
    {
        var commandType = assembly.GetType("ThesisEngine.Cli.AnalyzeCommand")!;
        var serviceType = assembly.GetType("ThesisEngine.Core.IDemoAnalysisService")!;
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
        var resultType = assembly.GetType("ThesisEngine.Core.AnalysisResult")!;
        var result = Activator.CreateInstance(resultType)!;
        var diagnosticType = assembly.GetType("ThesisEngine.Core.Diagnostic")!;
        var diagnosticArray = Array.CreateInstance(diagnosticType, diagnostics.Length);
        for (var index = 0; index < diagnostics.Length; index++)
            diagnosticArray.SetValue(diagnostics[index], index);
        resultType.GetProperty("Diagnostics")!.SetValue(result, diagnosticArray);
        resultType.GetProperty("RuleVersion")!.SetValue(result, "demo-0.1");
        resultType.GetProperty("Coverage")!.SetValue(result, coverage);
        resultType.GetProperty("UnsupportedScopes")!.SetValue(result, unsupported);
        return result;
    }

    private static object NewDiagnostic(Assembly assembly, string severity = "warning")
    {
        var diagnosticType = assembly.GetType("ThesisEngine.Core.Diagnostic")!;
        var diagnostic = Activator.CreateInstance(diagnosticType)!;
        diagnosticType.GetProperty("RuleId")!.SetValue(diagnostic, "demo/font-size");
        diagnosticType.GetProperty("Severity")!.SetValue(diagnostic, severity);
        diagnosticType.GetProperty("Location")!.SetValue(diagnostic, "paragraph[0]/run[0]");
        diagnosticType.GetProperty("Expected")!.SetValue(diagnostic, "12pt");
        diagnosticType.GetProperty("Actual")!.SetValue(diagnostic, "10.5pt");
        diagnosticType.GetProperty("FixLevel")!.SetValue(diagnostic, "manual");
        diagnosticType.GetProperty("Message")!.SetValue(diagnostic, "Font size differs from the demo target.");
        return diagnostic;
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
