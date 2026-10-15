using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

using System.Text.Json;
using System.Text.Json.Serialization;

namespace ThesisAddIn.Protocol;

public static class EngineProtocol
{
    public const string Version = "1.0";
    public const string AnalyzeOperation = "analyze";

    public const string InvalidRequest = "invalid_request";
    public const string ProtocolVersionMismatch = "protocol_version_mismatch";
    public const string ProtocolError = "protocol_error";
    public const string EngineError = "engine_error";
    public const string EngineExit = "engine_exit";
    public const string Timeout = "timeout";
    public const string Cancelled = "cancelled";
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

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public EngineError? Error { get; set; }
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

    public EngineResponse Analyze(string inputPath) => AnalyzeAsync(inputPath).GetAwaiter().GetResult();

    public async Task<EngineResponse> AnalyzeAsync(string inputPath, CancellationToken cancellationToken = default)
    {
        var request = new EngineRequest
        {
            RequestId = Guid.NewGuid().ToString("N"),
            Operation = EngineProtocol.AnalyzeOperation,
            InputPath = inputPath ?? string.Empty
        };

        if (string.IsNullOrWhiteSpace(inputPath) ||
            !string.Equals(Path.GetExtension(inputPath), ".docx", StringComparison.OrdinalIgnoreCase))
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
            if (response.Ok && response.Result is null)
            {
                CloseInputAndWait(process);
                return Failure(request.RequestId, EngineProtocol.ProtocolError, "成功响应缺少 result。", request);
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
        catch (Exception exception) when (exception is InvalidOperationException or IOException or UnauthorizedAccessException)
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
