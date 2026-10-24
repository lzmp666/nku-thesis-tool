using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using System.Text.Json;
using System.Text.Json.Serialization;

namespace ThesisAddIn.Protocol;

public static class EngineProtocol
{
    public const string Version = "1.2";
    public const string AnalyzeOperation = "analyze";
    public const string PlanOperation = "plan";
    public const string ApplyOperation = "apply";
    public const string AnnotateOperation = "annotate";

    public const string InvalidRequest = "invalid_request";
    public const string ProtocolVersionMismatch = "protocol_version_mismatch";
    public const string ProtocolError = "protocol_error";
    public const string EngineError = "engine_error";
    public const string EngineExit = "engine_exit";
    public const string Timeout = "timeout";
    public const string Cancelled = "cancelled";
    public const string PatchConflict = "patch_conflict";
}

public sealed class EngineRequest
{
    [JsonPropertyName("protocolVersion")]
    public string ProtocolVersion { get; set; } = EngineProtocol.Version;

    [JsonPropertyName("requestId")]
    public string RequestId { get; set; } = string.Empty;

    [JsonPropertyName("operation")]
    public string Operation { get; set; } = EngineProtocol.AnalyzeOperation;

    [JsonPropertyName("inputPath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>apply/annotate 操作的目标路径；必须与 inputPath 不同，且不得已存在。</summary>
    [JsonPropertyName("outputPath")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? OutputPath { get; set; }

    /// <summary>可选的规则编号过滤，用于“只处理当前这一项”。</summary>
    [JsonPropertyName("ruleId")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RuleId { get; set; }

    /// <summary>可选的位置过滤，与 ruleId 一起确定唯一的一条诊断。</summary>
    [JsonPropertyName("location")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Location { get; set; }

    /// <summary>所有属于当前合并问题的位置；旧客户端仍可只传 location。</summary>
    [JsonPropertyName("locations")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? Locations { get; set; }
}

public sealed class EngineResponse
{
    [JsonPropertyName("protocolVersion")]
    public string ProtocolVersion { get; set; } = EngineProtocol.Version;

    [JsonPropertyName("requestId")]
    public string RequestId { get; set; } = string.Empty;

    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("result")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public EngineAnalysisResult? Result { get; set; }

    /// <summary>plan 操作的返回：将要修改的内容与保留给用户的诊断。</summary>
    [JsonPropertyName("plan")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public EnginePatchPlan? Plan { get; set; }

    /// <summary>apply 操作的返回：已应用、失败和未处理的补丁。</summary>
    [JsonPropertyName("apply")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public EngineApplyResult? Apply { get; set; }

    /// <summary>annotate 操作写入的新文档及批注数量。</summary>
    [JsonPropertyName("annotation")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public EngineAnnotationResult? Annotation { get; set; }

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public EngineError? Error { get; set; }
}

public sealed class EngineAnnotationResult
{
    [JsonPropertyName("outputPath")]
    public string OutputPath { get; set; } = string.Empty;

    [JsonPropertyName("commentCount")]
    public int CommentCount { get; set; }
}

/// <summary>一条待执行的修复：位置、改前值、目标值和依据。</summary>
public sealed class EnginePatch
{
    [JsonPropertyName("ruleId")]
    public string RuleId { get; set; } = string.Empty;

    [JsonPropertyName("location")]
    public string Location { get; set; } = string.Empty;

    [JsonPropertyName("expectedBefore")]
    public string ExpectedBefore { get; set; } = string.Empty;

    [JsonPropertyName("desiredAfter")]
    public string DesiredAfter { get; set; } = string.Empty;

    [JsonPropertyName("reason")]
    public string Reason { get; set; } = string.Empty;
}

public sealed class EnginePatchFailure
{
    [JsonPropertyName("patch")]
    public EnginePatch Patch { get; set; } = new();

    [JsonPropertyName("reason")]
    public string Reason { get; set; } = string.Empty;
}

/// <summary>
/// 修复计划：Patches 只包含处理分类为 auto 的项目；Deferred 是保留给用户处理的诊断。
/// </summary>
public sealed class EnginePatchPlan
{
    [JsonPropertyName("patches")]
    public List<EnginePatch> Patches { get; set; } = new();

    [JsonPropertyName("deferred")]
    public List<EngineDiagnostic> Deferred { get; set; } = new();

    [JsonPropertyName("ruleVersion")]
    public string RuleVersion { get; set; } = string.Empty;

    [JsonPropertyName("coverage")]
    public string Coverage { get; set; } = string.Empty;

    [JsonPropertyName("unsupportedScopes")]
    public List<string> UnsupportedScopes { get; set; } = new();
}

public sealed class EngineApplyResult
{
    [JsonPropertyName("outputPath")]
    public string OutputPath { get; set; } = string.Empty;

    [JsonPropertyName("applied")]
    public List<EnginePatch> Applied { get; set; } = new();

    [JsonPropertyName("failures")]
    public List<EnginePatchFailure> Failures { get; set; } = new();

    [JsonPropertyName("unresolved")]
    public List<EnginePatch> Unresolved { get; set; } = new();
}

public sealed class EngineError
{
    [JsonPropertyName("code")]
    public string Code { get; set; } = string.Empty;

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;
}

public sealed class EngineAnalysisResult
{
    [JsonPropertyName("diagnostics")]
    public List<EngineDiagnostic> Diagnostics { get; set; } = new();

    [JsonPropertyName("ruleVersion")]
    public string RuleVersion { get; set; } = string.Empty;

    [JsonPropertyName("coverage")]
    public string Coverage { get; set; } = string.Empty;

    [JsonPropertyName("unsupportedScopes")]
    public List<string> UnsupportedScopes { get; set; } = new();
}

public sealed class EngineDiagnostic
{
    [JsonPropertyName("ruleId")]
    public string RuleId { get; set; } = string.Empty;

    [JsonPropertyName("severity")]
    public string Severity { get; set; } = string.Empty;

    [JsonPropertyName("location")]
    public string Location { get; set; } = string.Empty;

    [JsonPropertyName("expected")]
    public string Expected { get; set; } = string.Empty;

    [JsonPropertyName("actual")]
    public string Actual { get; set; } = string.Empty;

    [JsonPropertyName("fixLevel")]
    public string FixLevel { get; set; } = string.Empty;

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("occurrences")]
    public List<EngineDiagnosticOccurrence> Occurrences { get; set; } = new();
}

public sealed class EngineDiagnosticOccurrence
{
    [JsonPropertyName("location")]
    public string Location { get; set; } = string.Empty;

    [JsonPropertyName("actual")]
    public string Actual { get; set; } = string.Empty;
}

public static class EngineProtocolCodec
{
    private static readonly JsonSerializerOptions Options = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    public static string SerializeRequest(EngineRequest request) => Serialize(request);

    public static string SerializeResponse(EngineResponse response) => Serialize(response);

    public static bool TryDeserializeRequest(string line, out EngineRequest? request, out string error)
    {
        request = null;
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(line))
        {
            error = "协议输入为空。";
            return false;
        }

        try
        {
            request = JsonSerializer.Deserialize<EngineRequest>(line, Options);
            if (request is null)
            {
                error = "协议输入不是请求对象。";
                return false;
            }

            return true;
        }
        catch (JsonException exception)
        {
            error = $"协议 JSON 无效：{exception.Message}";
            return false;
        }
    }

    public static bool TryDeserializeResponse(string line, out EngineResponse? response, out string error)
    {
        response = null;
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(line))
        {
            error = "引擎没有返回协议响应。";
            return false;
        }

        try
        {
            response = JsonSerializer.Deserialize<EngineResponse>(line, Options);
            if (response is null)
            {
                error = "引擎响应不是响应对象。";
                return false;
            }

            return true;
        }
        catch (JsonException exception)
        {
            error = $"引擎响应 JSON 无效：{exception.Message}";
            return false;
        }
    }

    private static string Serialize<T>(T value) => JsonSerializer.Serialize(value, Options);
}

public sealed class EngineProcessOptions
{
    public EngineProcessOptions(string fileName, string arguments, TimeSpan timeout)
    {
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("必须提供进程文件名。", nameof(fileName));
        if (timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(timeout));
        FileName = fileName;
        Arguments = arguments ?? string.Empty;
        Timeout = timeout;
    }

    public string FileName { get; }
    public string Arguments { get; }
    public string? WorkingDirectory { get; set; }
    public TimeSpan Timeout { get; }

    public static EngineProcessOptions ForDotnet(string hostAssemblyPath, TimeSpan? timeout = null)
    {
        if (string.IsNullOrWhiteSpace(hostAssemblyPath)) throw new ArgumentException("必须提供协议宿主路径。", nameof(hostAssemblyPath));
        return new EngineProcessOptions("dotnet", Quote(hostAssemblyPath) + " --protocol-stdio", timeout ?? TimeSpan.FromSeconds(30));
    }

    private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";
}

public sealed class EngineProtocolClient
{
    private readonly EngineProcessOptions _options;

    public EngineProtocolClient(EngineProcessOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    public EngineResponse Analyze(string inputPath) =>
        Send(CreateRequest(EngineProtocol.AnalyzeOperation, inputPath, null, null, null));

    public EngineResponse Plan(string inputPath, string? ruleId, string? location) =>
        Send(CreateRequest(EngineProtocol.PlanOperation, inputPath, null, ruleId, location));

    public EngineResponse Plan(string inputPath, string? ruleId, string? location, IReadOnlyList<string>? locations) =>
        Send(CreateRequest(EngineProtocol.PlanOperation, inputPath, null, ruleId, location, locations));

    public Task<EngineResponse> PlanAsync(
        string inputPath,
        string? ruleId,
        string? location,
        CancellationToken cancellationToken = default) =>
        SendAsync(CreateRequest(EngineProtocol.PlanOperation, inputPath, null, ruleId, location), cancellationToken);

    public Task<EngineResponse> PlanAsync(
        string inputPath,
        string? ruleId,
        string? location,
        IReadOnlyList<string>? locations,
        CancellationToken cancellationToken = default) =>
        SendAsync(CreateRequest(EngineProtocol.PlanOperation, inputPath, null, ruleId, location, locations), cancellationToken);

    public EngineResponse Apply(string inputPath, string outputPath, string? ruleId, string? location) =>
        Send(CreateRequest(EngineProtocol.ApplyOperation, inputPath, outputPath, ruleId, location));

    public EngineResponse Apply(string inputPath, string outputPath, string? ruleId, string? location, IReadOnlyList<string>? locations) =>
        Send(CreateRequest(EngineProtocol.ApplyOperation, inputPath, outputPath, ruleId, location, locations));

    public EngineResponse Annotate(string inputPath, string outputPath) =>
        Send(CreateRequest(EngineProtocol.AnnotateOperation, inputPath, outputPath, null, null));

    public Task<EngineResponse> AnnotateAsync(
        string inputPath,
        string outputPath,
        CancellationToken cancellationToken = default) =>
        SendAsync(CreateRequest(EngineProtocol.AnnotateOperation, inputPath, outputPath, null, null), cancellationToken);

    public Task<EngineResponse> ApplyAsync(
        string inputPath,
        string outputPath,
        string? ruleId,
        string? location,
        CancellationToken cancellationToken = default) =>
        SendAsync(CreateRequest(EngineProtocol.ApplyOperation, inputPath, outputPath, ruleId, location), cancellationToken);

    public Task<EngineResponse> ApplyAsync(
        string inputPath,
        string outputPath,
        string? ruleId,
        string? location,
        IReadOnlyList<string>? locations,
        CancellationToken cancellationToken = default) =>
        SendAsync(CreateRequest(EngineProtocol.ApplyOperation, inputPath, outputPath, ruleId, location, locations), cancellationToken);

    public Task<EngineResponse> AnalyzeAsync(string inputPath, CancellationToken cancellationToken = default) =>
        SendAsync(CreateRequest(EngineProtocol.AnalyzeOperation, inputPath, null, null, null), cancellationToken);

    public EngineResponse Send(EngineRequest request) => SendAsync(request).GetAwaiter().GetResult();

    private static EngineRequest CreateRequest(string operation, string inputPath, string? outputPath, string? ruleId, string? location,
        IReadOnlyList<string>? locations = null) => new()
    {
        RequestId = Guid.NewGuid().ToString("N"),
        Operation = operation,
        InputPath = inputPath ?? string.Empty,
        OutputPath = outputPath,
        RuleId = ruleId,
        Location = location,
        Locations = locations?.ToList()
    };

    public async Task<EngineResponse> SendAsync(EngineRequest request, CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        request.RequestId = string.IsNullOrWhiteSpace(request.RequestId) ? Guid.NewGuid().ToString("N") : request.RequestId;

        if (string.IsNullOrWhiteSpace(request.InputPath) ||
            !string.Equals(Path.GetExtension(request.InputPath), ".docx", StringComparison.OrdinalIgnoreCase))
            return Failure(request.RequestId, EngineProtocol.InvalidRequest, "输入文件必须是 .docx。", request);

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = _options.FileName,
                Arguments = _options.Arguments,
                WorkingDirectory = _options.WorkingDirectory ?? string.Empty,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            }
        };

        try
        {
            if (!process.Start())
                return Failure(request.RequestId, EngineProtocol.EngineExit, "无法启动引擎进程。", request);

            await process.StandardInput.WriteLineAsync(EngineProtocolCodec.SerializeRequest(request)).ConfigureAwait(false);
            await process.StandardInput.FlushAsync().ConfigureAwait(false);

            var responseTask = process.StandardOutput.ReadLineAsync();
            var delayTask = Task.Delay(_options.Timeout, cancellationToken);
            var completed = await Task.WhenAny(responseTask, delayTask).ConfigureAwait(false);
            if (completed != responseTask)
            {
                TryTerminate(process);
                return cancellationToken.IsCancellationRequested
                    ? Failure(request.RequestId, EngineProtocol.Cancelled, "调用已取消。", request)
                    : Failure(request.RequestId, EngineProtocol.Timeout, $"引擎在 {_options.Timeout.TotalSeconds:0.#} 秒内未返回结果。", request);
            }

            var line = await responseTask.ConfigureAwait(false);
            if (line is null)
            {
                return Failure(request.RequestId, EngineProtocol.EngineExit,
                    process.HasExited ? $"引擎进程已退出，退出码为 {process.ExitCode}。" : "引擎进程未返回响应。", request);
            }

            if (!EngineProtocolCodec.TryDeserializeResponse(line ?? string.Empty, out var response, out var error))
            {
                TryTerminate(process);
                return Failure(request.RequestId, EngineProtocol.ProtocolError, error, request);
            }

            if (!string.Equals(response!.ProtocolVersion, EngineProtocol.Version, StringComparison.Ordinal))
            {
                CloseInputAndWait(process);
                return Failure(request.RequestId, EngineProtocol.ProtocolVersionMismatch, "引擎协议版本不匹配。", request);
            }
            if (!string.Equals(response.RequestId, request.RequestId, StringComparison.Ordinal))
            {
                CloseInputAndWait(process);
                return Failure(request.RequestId, EngineProtocol.ProtocolError, "引擎响应 requestId 不匹配。", request);
            }
            if (response.Ok && !HasPayloadFor(request.Operation, response))
            {
                CloseInputAndWait(process);
                return Failure(request.RequestId, EngineProtocol.ProtocolError,
                    "成功响应缺少与操作匹配的结果（" + request.Operation + "）。", request);
            }
            if (!response.Ok && response.Error is null)
            {
                CloseInputAndWait(process);
                return Failure(request.RequestId, EngineProtocol.ProtocolError, "失败响应缺少 error。", request);
            }
            CloseInputAndWait(process);
            return response;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            TryTerminate(process);
            return Failure(request.RequestId, EngineProtocol.Cancelled, "调用已取消。", request);
        }
        catch (Exception exception) when (exception is InvalidOperationException or IOException or UnauthorizedAccessException or System.ComponentModel.Win32Exception)
        {
            TryTerminate(process);
            return Failure(request.RequestId, EngineProtocol.EngineExit, $"引擎进程启动或通信失败：{exception.Message}", request);
        }
    }

    private static EngineResponse Failure(string requestId, string code, string message, EngineRequest request) => new()
    {
        ProtocolVersion = EngineProtocol.Version,
        RequestId = requestId,
        Ok = false,
        Error = new EngineError { Code = code, Message = message }
    };

    /// <summary>
    /// 成功响应必须带上与操作匹配的载荷：analyze 要 result，plan 要 plan，apply 要 apply。
    /// 早期实现只检查 result，导致 plan/apply 的成功响应被误判为协议异常。
    /// </summary>
    private static bool HasPayloadFor(string operation, EngineResponse response) => operation switch
    {
        EngineProtocol.PlanOperation => response.Plan is not null,
        EngineProtocol.ApplyOperation => response.Apply is not null,
        EngineProtocol.AnnotateOperation => response.Annotation is not null && response.Result is not null,
        _ => response.Result is not null
    };

    private static void CloseInputAndWait(Process process)
    {
        try
        {
            process.StandardInput.Close();
            process.WaitForExit(1000);
        }
        catch (InvalidOperationException) { }
        catch (System.ComponentModel.Win32Exception) { }
    }

    private static void TryTerminate(Process process)
    {
        try
        {
            if (!process.HasExited) process.Kill();
        }
        catch (InvalidOperationException) { }
        catch (System.ComponentModel.Win32Exception) { }
    }
}

