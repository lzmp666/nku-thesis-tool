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
                    new()
                    {
                        RuleId = "font-size",
                        Location = "word/document.xml#paragraph[0]/run[0]",
                        Occurrences = new List<EngineDiagnosticOccurrence>
                        {
                            new() { Location = "word/document.xml#paragraph[0]/run[0]", Actual = "10pt" },
                            new() { Location = "word/document.xml#paragraph[1]/run[0]", Actual = "10pt" }
                        }
                    }
                }
            }
        };
        Assert.True(EngineProtocolCodec.TryDeserializeResponse(
            EngineProtocolCodec.SerializeResponse(response), out var decodedResponse, out var responseError), responseError);
        Assert.True(decodedResponse!.Ok);
        EngineDiagnostic decodedDiagnostic = Assert.Single(decodedResponse.Result!.Diagnostics);
        Assert.Equal("font-size", decodedDiagnostic.RuleId);
        Assert.Equal(2, decodedDiagnostic.Occurrences.Count);
        Assert.Equal("word/document.xml#paragraph[1]/run[0]", decodedDiagnostic.Occurrences[1].Location);

        var locationRequest = new EngineRequest
        {
            RequestId = "group-fix",
            Operation = EngineProtocol.PlanOperation,
            InputPath = @"C:\temp\sample.docx",
            RuleId = "font-size",
            Locations = new List<string>
            {
                "word/document.xml#paragraph[0]/run[0]",
                "word/document.xml#paragraph[1]/run[0]"
            }
        };
        Assert.True(EngineProtocolCodec.TryDeserializeRequest(
            EngineProtocolCodec.SerializeRequest(locationRequest), out var decodedLocationRequest, out var locationRequestError), locationRequestError);
        Assert.Equal(locationRequest.Locations, decodedLocationRequest!.Locations);
    }

    [Fact]
    public async Task ProtocolHostAnalyzesSyntheticDocumentThroughJsonLines()
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
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
        var host = FindProtocolHost(root);
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "clean_sample.docx");

        var client = new EngineProtocolClient(EngineProcessOptions.ForDotnet(host, TimeSpan.FromSeconds(5)));
        var response = await client.AnalyzeAsync(input);

        Assert.True(response.Ok, response.Error?.Message);
        Assert.Empty(response.Result!.Diagnostics);
        Assert.Equal("demo-0.1.0", response.Result.RuleVersion);
    }

    [Fact]
    public async Task ProtocolHostAnnotatesNewCopyAndRefreshesToolCommentsIdempotently()
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");
        var firstOutput = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var secondOutput = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var originalHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(input)));
        try
        {
            var first = await RunHost(host, new EngineRequest
            {
                RequestId = "annotate-1",
                Operation = EngineProtocol.AnnotateOperation,
                InputPath = input,
                OutputPath = firstOutput
            });
            Assert.True(first.Ok, first.Error?.Message);
            Assert.Equal(firstOutput, first.Annotation!.OutputPath);
            Assert.Equal(3, first.Annotation.CommentCount);
            Assert.Equal(3, first.Result!.Diagnostics.Count);

            var second = await RunHost(host, new EngineRequest
            {
                RequestId = "annotate-2",
                Operation = EngineProtocol.AnnotateOperation,
                InputPath = firstOutput,
                OutputPath = secondOutput
            });
            Assert.True(second.Ok, second.Error?.Message);
            Assert.Equal(3, second.Annotation!.CommentCount);
            Assert.Equal(first.Result.Diagnostics.Select(item => item.Location), second.Result!.Diagnostics.Select(item => item.Location));
            Assert.Equal(originalHash, Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(input))));

            using var archive = System.IO.Compression.ZipFile.OpenRead(secondOutput);
            System.Xml.Linq.XDocument comments;
            using (var stream = archive.GetEntry("word/comments.xml")!.Open()) comments = System.Xml.Linq.XDocument.Load(stream);
            var w = System.Xml.Linq.XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var allCommentText = string.Concat(comments.Descendants(w + "t").Select(item => item.Value));
            Assert.Equal(3, comments.Root!.Elements(w + "comment").Count());
            Assert.Equal(3, CountOccurrences(allCommentText, "[NKU-TOOL-ANNOTATION:v1]"));
        }
        finally
        {
            if (File.Exists(firstOutput)) File.Delete(firstOutput);
            if (File.Exists(secondOutput)) File.Delete(secondOutput);
        }
    }

    [Fact]
    public async Task ProtocolHostBatchRepairsAndReanalyzesFullSyntheticThesisWithoutChangingItsText()
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_complete_thesis.docx");
        var appliedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "-full-thesis-fixed.docx");
        var annotatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "-full-thesis-reviewed.docx");
        var originalHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(input)));
        var originalText = ReadBodyText(input);
        try
        {
            var client = new EngineProtocolClient(EngineProcessOptions.ForDotnet(host, TimeSpan.FromMinutes(2)));
            var plan = await client.PlanAsync(input, null, null);
            Assert.True(plan.Ok, plan.Error?.Message);
            Assert.Equal(3, plan.Plan!.Patches.Count);

            var apply = await client.ApplyAsync(input, appliedPath, null, null);
            Assert.True(apply.Ok, apply.Error?.Message);
            Assert.NotEmpty(apply.Apply!.Applied);
            Assert.Empty(apply.Apply.Failures);
            Assert.Equal(originalText, ReadBodyText(appliedPath));

            var refreshed = await client.AnnotateAsync(appliedPath, annotatedPath);
            Assert.True(refreshed.Ok, refreshed.Error?.Message);
            Assert.NotNull(refreshed.Annotation);
            Assert.NotNull(refreshed.Result);
            Assert.Empty(refreshed.Result!.Diagnostics);
            Assert.Equal(originalText, ReadBodyText(annotatedPath));
            Assert.Equal(originalHash, Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(input))));
        }
        finally
        {
            if (File.Exists(appliedPath)) File.Delete(appliedPath);
            if (File.Exists(annotatedPath)) File.Delete(annotatedPath);
        }
    }

    [Fact]
    public async Task ProtocolHostAnnotateRejectsExistingOutputWithoutOverwritingIt()
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "clean_sample.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        File.WriteAllText(output, "keep this file");
        try
        {
            var response = await RunHost(host, new EngineRequest
            {
                Operation = EngineProtocol.AnnotateOperation,
                InputPath = input,
                OutputPath = output
            });
            Assert.False(response.Ok);
            Assert.Equal(EngineProtocol.PatchConflict, response.Error!.Code);
            Assert.Equal("keep this file", File.ReadAllText(output));
        }
        finally
        {
            File.Delete(output);
        }
    }

    [Theory]
    [InlineData("1.1")]
    [InlineData("9.0")]
    public async Task ProtocolHostRejectsAnUnsupportedProtocolVersion(string version)
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");

        var response = await RunHost(host, new EngineRequest
        {
            ProtocolVersion = version,
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

    [Fact]
    public async Task ClientReportsMissingExecutableAsEngineExit()
    {
        var root = FindRepositoryRoot();
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "clean_sample.docx");
        Assert.True(File.Exists(input), $"Synthetic input not found: {input}");

        var missingExecutable = Path.Combine(
            Path.GetTempPath(),
            $"nku-thesis-missing-engine-{Guid.NewGuid():N}.exe");
        Assert.False(File.Exists(missingExecutable));

        var client = new EngineProtocolClient(
            new EngineProcessOptions(missingExecutable, string.Empty, TimeSpan.FromSeconds(1)));
        var response = await client.AnalyzeAsync(input);

        Assert.False(response.Ok);
        Assert.Equal(EngineProtocol.EngineExit, response.Error!.Code);
        Assert.Contains("引擎进程启动或通信失败", response.Error.Message);
    }

    [Fact]
    public async Task ProtocolHostPlansAutoPatchesAndKeepsDeferredDiagnostics()
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");

        var response = await RunHost(host, new EngineRequest
        {
            RequestId = "plan-1",
            Operation = EngineProtocol.PlanOperation,
            InputPath = input
        });

        Assert.True(response.Ok, response.Error?.Message);
        Assert.Equal("plan-1", response.RequestId);
        Assert.Null(response.Result);
        Assert.NotNull(response.Plan);
        Assert.Equal(3, response.Plan!.Patches.Count);
        Assert.Contains(response.Plan.Patches, patch => patch.RuleId == "demo/east-asian-font");
        Assert.Contains(response.Plan.Patches, patch => patch.RuleId == "demo/western-font");
        Assert.Contains(response.Plan.Patches, patch => patch.RuleId == "demo/font-size");
        Assert.Empty(response.Plan.Deferred);
    }

    [Fact]
    public async Task ProtocolHostAppliesPatchesToNewFileAndReportsRecheck()
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "-fixed.docx");
        var before = File.ReadAllBytes(input);
        try
        {
            var apply = await RunHost(host, new EngineRequest
            {
                RequestId = "apply-1",
                Operation = EngineProtocol.ApplyOperation,
                InputPath = input,
                OutputPath = output
            });

            Assert.True(apply.Ok, apply.Error?.Message);
            Assert.NotNull(apply.Apply);
            Assert.Equal(3, apply.Apply!.Applied.Count);
            Assert.Empty(apply.Apply.Failures);
            Assert.Equal(Path.GetFullPath(output), apply.Apply.OutputPath);
            Assert.True(File.Exists(output));
            Assert.Equal(before, File.ReadAllBytes(input));

            var recheck = await RunHost(host, new EngineRequest { RequestId = "recheck-1", InputPath = output });
            Assert.True(recheck.Ok, recheck.Error?.Message);
            Assert.Empty(recheck.Result!.Diagnostics);

            var replan = await RunHost(host, new EngineRequest
            {
                RequestId = "replan-1",
                Operation = EngineProtocol.PlanOperation,
                InputPath = output
            });
            Assert.True(replan.Ok, replan.Error?.Message);
            Assert.Empty(replan.Plan!.Patches);
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public async Task ClientAcceptsPlanAndApplyPayloadsFromHost()
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "-client-fixed.docx");
        try
        {
            var client = new EngineProtocolClient(EngineProcessOptions.ForDotnet(host, TimeSpan.FromSeconds(15)));

            // plan 的成功响应带的是 plan 载荷；早期客户端只校验 result，会把它误判成协议异常。
            var plan = await client.PlanAsync(input, null, null);
            Assert.True(plan.Ok, plan.Error?.Message);
            Assert.NotNull(plan.Plan);
            Assert.Equal(3, plan.Plan!.Patches.Count);

            var filtered = await client.PlanAsync(input, "demo/font-size", "word/document.xml#paragraph[3]/run[0]");
            Assert.True(filtered.Ok, filtered.Error?.Message);
            var single = Assert.Single(filtered.Plan!.Patches);
            Assert.Equal("demo/font-size", single.RuleId);
            // 过滤后保留被排除的中文/西文字体修复项，当前字号项只计划一次。
            Assert.Equal(2, filtered.Plan.Deferred.Count);
            Assert.Contains(filtered.Plan.Deferred, item => item.RuleId == "demo/western-font");
            Assert.Contains(filtered.Plan.Deferred, item => item.RuleId == "demo/east-asian-font");

            var apply = await client.ApplyAsync(input, output, null, null);
            Assert.True(apply.Ok, apply.Error?.Message);
            Assert.NotNull(apply.Apply);
            Assert.Equal(3, apply.Apply!.Applied.Count);
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public async Task ProtocolHostRefusesToOverwriteExistingOutput()
    {
        var root = FindRepositoryRoot();
        var host = FindProtocolHost(root);
        Assert.True(File.Exists(host), $"Protocol host was not built: {host}");
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "-existing.docx");
        File.WriteAllText(output, "existing");
        try
        {
            var response = await RunHost(host, new EngineRequest
            {
                RequestId = "overwrite-1",
                Operation = EngineProtocol.ApplyOperation,
                InputPath = input,
                OutputPath = output
            });

            Assert.False(response.Ok);
            Assert.Equal(EngineProtocol.PatchConflict, response.Error!.Code);
            Assert.Equal("existing", File.ReadAllText(output));
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
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

    private static string ReadBodyText(string path)
    {
        using var archive = System.IO.Compression.ZipFile.OpenRead(path);
        using var stream = archive.GetEntry("word/document.xml")!.Open();
        var document = System.Xml.Linq.XDocument.Load(stream);
        var word = System.Xml.Linq.XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        return string.Concat(document.Descendants(word + "t").Select(element => element.Value));
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static string FindProtocolHost(string root)
    {
        var projectOutput = Path.Combine(root, "src", "ThesisAddIn.VSTO", "ProtocolHost", "bin");
        var configuration = AppContext.BaseDirectory.Contains(
            Path.DirectorySeparatorChar + "Release" + Path.DirectorySeparatorChar,
            StringComparison.OrdinalIgnoreCase)
            ? "Release"
            : "Debug";
        var configuredPath = Path.Combine(projectOutput, configuration, "net10.0", "ThesisEngine.ProtocolHost.dll");
        if (File.Exists(configuredPath)) return configuredPath;

        var fallback = Path.Combine(projectOutput, configuration == "Release" ? "Debug" : "Release", "net10.0", "ThesisEngine.ProtocolHost.dll");
        return File.Exists(fallback) ? fallback : configuredPath;
    }

    private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";

    private static int CountOccurrences(string value, string token)
    {
        var count = 0;
        for (var index = 0; (index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0; index += token.Length)
            count++;
        return count;
    }
}

