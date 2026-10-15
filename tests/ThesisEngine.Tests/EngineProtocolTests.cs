using System.Diagnostics;
using ThesisAddIn.Protocol;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class EngineProtocolTests
{
    [Fact]
    public void RequestAndResponseRoundTripPreservesVersionedFields()
    {
        var request = new EngineRequest
        {
            RequestId = "request-1",
            Operation = EngineProtocol.AnalyzeOperation,
            InputPath = @"C:\temp\sample.docx"
        };

        Assert.True(EngineProtocolCodec.TryDeserializeRequest(
            EngineProtocolCodec.SerializeRequest(request), out var decodedRequest, out var requestError), requestError);
        Assert.Equal(request.RequestId, decodedRequest!.RequestId);
        Assert.Equal(request.InputPath, decodedRequest.InputPath);

        var response = new EngineResponse
        {
            RequestId = request.RequestId,
            Ok = true,
            Result = new EngineAnalysisResult
            {
                RuleVersion = "demo-0.1.0",
                Coverage = "正文 run",
                Diagnostics = new List<EngineDiagnostic>
                {
                    new() { RuleId = "font-size", Location = "word/document.xml#paragraph[0]/run[0]" }
                }
            }
        };
        Assert.True(EngineProtocolCodec.TryDeserializeResponse(
            EngineProtocolCodec.SerializeResponse(response), out var decodedResponse, out var responseError), responseError);
        Assert.True(decodedResponse!.Ok);
        Assert.Equal("font-size", Assert.Single(decodedResponse.Result!.Diagnostics).RuleId);
    }

    [Fact]
    public async Task ProtocolHostAnalyzesSyntheticDocumentThroughJsonLines()
    {
        var root = FindRepositoryRoot();
        var host = Path.Combine(root, "src", "ThesisAddIn.VSTO", "ProtocolHost", "bin", "Debug", "net10.0", "ThesisEngine.ProtocolHost.dll");
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");

        var result = await RunHost(host, new EngineRequest
        {
            RequestId = "integration-1",
            InputPath = input
        });

        Assert.True(result.Ok, result.Error?.Message);
        Assert.Equal("integration-1", result.RequestId);
        Assert.Equal("demo-0.1.0", result.Result!.RuleVersion);
        Assert.Equal(3, result.Result.Diagnostics.Count);
    }

    [Fact]
    public async Task ClientInvokesProtocolHostAndReturnsStructuredAnalysis()
    {
        var root = FindRepositoryRoot();
        var host = Path.Combine(root, "src", "ThesisAddIn.VSTO", "ProtocolHost", "bin", "Debug", "net10.0", "ThesisEngine.ProtocolHost.dll");
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "clean_sample.docx");

        var client = new EngineProtocolClient(EngineProcessOptions.ForDotnet(host, TimeSpan.FromSeconds(5)));
        var response = await client.AnalyzeAsync(input);

        Assert.True(response.Ok, response.Error?.Message);
        Assert.Empty(response.Result!.Diagnostics);
        Assert.Equal("demo-0.1.0", response.Result.RuleVersion);
    }

    [Fact]
    public async Task ProtocolHostRejectsAnUnknownProtocolVersion()
    {
        var root = FindRepositoryRoot();
        var host = Path.Combine(root, "src", "ThesisAddIn.VSTO", "ProtocolHost", "bin", "Debug", "net10.0", "ThesisEngine.ProtocolHost.dll");
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");

        var response = await RunHost(host, new EngineRequest
        {
            ProtocolVersion = "9.0",
            RequestId = "version-1",
            InputPath = "not-used.docx"
        });

        Assert.False(response.Ok);
        Assert.Equal(EngineProtocol.ProtocolVersionMismatch, response.Error!.Code);
    }

    [Fact]
    public async Task ClientReportsTimeoutAndTerminatesHungEngine()
    {
        if (OperatingSystem.IsWindows()) return;

        var input = Path.Combine(Path.GetTempPath(), "protocol-timeout.docx");
        var client = new EngineProtocolClient(new EngineProcessOptions("/bin/sh", "-c \"sleep 2\"", TimeSpan.FromMilliseconds(100)));
        var response = await client.AnalyzeAsync(input);

        Assert.False(response.Ok);
        Assert.Equal(EngineProtocol.Timeout, response.Error!.Code);
    }

    [Fact]
    public async Task ClientReportsAnEngineThatExitsBeforeResponding()
    {
        if (OperatingSystem.IsWindows()) return;

        var client = new EngineProtocolClient(new EngineProcessOptions("/bin/sh", "-c \"exit 7\"", TimeSpan.FromSeconds(1)));
        var response = await client.AnalyzeAsync("engine-exit.docx");

        Assert.False(response.Ok);
        Assert.Equal(EngineProtocol.EngineExit, response.Error!.Code);
    }

    [Fact]
    public async Task ClientRejectsUnsupportedInputWithoutStartingEngine()
    {
        var client = new EngineProtocolClient(new EngineProcessOptions("does-not-exist", string.Empty, TimeSpan.FromSeconds(1)));
        var response = await client.AnalyzeAsync("readme.txt");

        Assert.False(response.Ok);
        Assert.Equal(EngineProtocol.InvalidRequest, response.Error!.Code);
    }

    private static async Task<EngineResponse> RunHost(string host, EngineRequest request)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = Quote(host) + " --protocol-stdio",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            }
        };
        Assert.True(process.Start());
        await process.StandardInput.WriteLineAsync(EngineProtocolCodec.SerializeRequest(request));
        process.StandardInput.Close();
        var line = await process.StandardOutput.ReadLineAsync();
        Assert.True(EngineProtocolCodec.TryDeserializeResponse(line ?? string.Empty, out var response, out var error), error);
        return response!;
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";
}
