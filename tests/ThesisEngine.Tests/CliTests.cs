using System.Diagnostics;
using System.IO.Compression;
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

    private static async Task<(int ExitCode, string Output, string Error)> RunCli(params string[] args)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln")))
            root = root.Parent;
        Assert.NotNull(root);
        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
        var cli = Path.Combine(root.FullName, "src", "ThesisEngine.Cli", "bin", configuration, "net10.0", "ThesisEngine.Cli.dll");
        Assert.True(File.Exists(cli), $"CLI was not built: {cli}");
        var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
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
