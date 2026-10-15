using ThesisAddIn.Protocol;
using ThesisEngine.Core;

if (args.Length != 1 || !string.Equals(args[0], "--protocol-stdio", StringComparison.Ordinal))
{
    Console.Error.WriteLine("用法：ThesisEngine.ProtocolHost --protocol-stdio");
    return 2;
}

Console.InputEncoding = System.Text.Encoding.UTF8;
Console.OutputEncoding = System.Text.Encoding.UTF8;
var service = new DemoAnalysisService(new WordFormattingReader(), new FontSizeAnalyzer());

string? line;
while ((line = Console.ReadLine()) is not null)
{
    if (!EngineProtocolCodec.TryDeserializeRequest(line, out var request, out var parseError))
    {
        Write(ProtocolHostResponse.Error("unknown", EngineProtocol.ProtocolError, parseError));
        continue;
    }

    if (!string.Equals(request!.ProtocolVersion, EngineProtocol.Version, StringComparison.Ordinal))
    {
        Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.ProtocolVersionMismatch, "协议版本不匹配。"));
        continue;
    }

    if (!string.Equals(request.Operation, EngineProtocol.AnalyzeOperation, StringComparison.Ordinal))
    {
        Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "不支持的协议操作。"));
        continue;
    }

    if (string.IsNullOrWhiteSpace(request.InputPath) ||
        !string.Equals(Path.GetExtension(request.InputPath), ".docx", StringComparison.OrdinalIgnoreCase))
    {
        Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "输入文件必须是 .docx。"));
        continue;
    }

    try
    {
        var fullPath = Path.GetFullPath(request.InputPath);
        if (!File.Exists(fullPath))
        {
            Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "找不到输入文件。"));
            continue;
        }

        var analysis = service.Analyze(fullPath, DemoRuleSet.NkuDraftSample);
        Write(ProtocolHostResponse.Success(request.RequestId, analysis));
    }
    catch (Exception exception) when (exception is ArgumentException or InvalidDataException or IOException or UnauthorizedAccessException or NotSupportedException or System.Xml.XmlException)
    {
        Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.EngineError, $"引擎分析失败：{exception.Message}"));
    }
}

return 0;

static void Write(EngineResponse response)
{
    Console.WriteLine(EngineProtocolCodec.SerializeResponse(response));
    Console.Out.Flush();
}

static class ProtocolHostResponse
{
    public static EngineResponse Success(string requestId, AnalysisResult result) => new()
    {
        ProtocolVersion = EngineProtocol.Version,
        RequestId = requestId,
        Ok = true,
        Result = new EngineAnalysisResult
        {
            RuleVersion = result.RuleVersion,
            Coverage = result.Coverage,
            UnsupportedScopes = result.UnsupportedScopes?.ToList() ?? new List<string>(),
            Diagnostics = (result.Diagnostics ?? Array.Empty<Diagnostic>()).Select(d => new EngineDiagnostic
            {
                RuleId = d.RuleId,
                Severity = d.Severity,
                Location = d.Location,
                Expected = d.Expected,
                Actual = d.Actual,
                FixLevel = d.FixLevel,
                Message = d.Message
            }).ToList()
        }
    };

    public static EngineResponse Error(string requestId, string code, string message) => new()
    {
        ProtocolVersion = EngineProtocol.Version,
        RequestId = requestId,
        Ok = false,
        Error = new EngineError { Code = code, Message = message }
    };
}
