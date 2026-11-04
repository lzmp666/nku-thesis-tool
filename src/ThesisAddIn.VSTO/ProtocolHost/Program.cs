using ThesisAddIn.Protocol;
using ThesisEngine.Core;

if (args.Length != 1 || !string.Equals(args[0], "--protocol-stdio", StringComparison.Ordinal))
{
    Console.Error.WriteLine("用法：ThesisEngine.ProtocolHost --protocol-stdio");
    return 2;
}

Console.InputEncoding = System.Text.Encoding.UTF8;
Console.OutputEncoding = System.Text.Encoding.UTF8;
var service = new DemoAnalysisService(new WordFormattingReader(inferUnstyledStructure: true), new FontSizeAnalyzer());

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

    if (!TryResolveInput(request, out var fullPath, out var inputError))
    {
        Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, inputError));
        continue;
    }

    try
    {
        switch (request.Operation)
        {
            case EngineProtocol.AnalyzeOperation:
                Write(ProtocolHostResponse.Analysis(request.RequestId, Analyze(fullPath)));
                break;
            case EngineProtocol.PlanOperation:
                Write(ProtocolHostResponse.Plan(request.RequestId, BuildPlan(fullPath, request)));
                break;
            case EngineProtocol.ApplyOperation:
                Write(ApplyPlan(request, fullPath));
                break;
            case EngineProtocol.AnnotateOperation:
                Write(Annotate(request, fullPath));
                break;
            case EngineProtocol.StructureReviewOperation:
            case EngineProtocol.StructurePlanOperation:
            case EngineProtocol.StructureVerifyOperation:
                Write(StructureProtocolBridge.Handle(request, fullPath));
                break;
            default:
                Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "不支持的协议操作。"));
                break;
        }
    }
    catch (ArgumentException exception)
    {
        Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "请求无效：" + exception.Message));
    }
    catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException
        or NotSupportedException or System.Xml.XmlException)
    {
        Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.EngineError, "引擎处理失败：" + exception.Message));
    }
    catch (Exception exception)
    {
        Write(ProtocolHostResponse.Error(request.RequestId, EngineProtocol.PatchConflict, "修复未能执行：" + exception.Message));
    }
}

return 0;

AnalysisResult Analyze(string fullPath) => service.Analyze(fullPath, DemoRuleSet.NkuDraftSample);

EnginePatchPlan BuildPlan(string fullPath, EngineRequest request)
{
    var analysis = Analyze(fullPath);
    var plan = PatchEngine.PlanPatches(analysis);
    var selected = SelectPatches(plan, request);
    return ProtocolHostResponse.ToPlanTransfer(plan, selected, analysis, request);
}

EngineResponse ApplyPlan(EngineRequest request, string fullPath)
{
    if (string.IsNullOrWhiteSpace(request.OutputPath))
        return ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "apply 操作必须提供 outputPath。");
    if (!string.Equals(Path.GetExtension(request.OutputPath), ".docx", StringComparison.OrdinalIgnoreCase))
        return ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "输出文件必须是 .docx。");

    var analysis = Analyze(fullPath);
    var plan = PatchEngine.PlanPatches(analysis);
    var selected = SelectPatches(plan, request);
    if (selected.Count == 0)
        return ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest,
            "没有可自动修复的项目；请先查看修复计划或改为人工处理。");

    var outputPath = Path.GetFullPath(request.OutputPath);
    if (File.Exists(outputPath))
        return ProtocolHostResponse.Error(request.RequestId, EngineProtocol.PatchConflict,
            "输出文件已存在：" + outputPath + "。为避免覆盖已有文件，请换一个文件名。");

    var result = PatchEngine.ApplyPatches(fullPath, selected, outputPath);
    return ProtocolHostResponse.Applied(request.RequestId, result);
}

EngineResponse Annotate(EngineRequest request, string fullPath)
{
    if (string.IsNullOrWhiteSpace(request.OutputPath))
        return ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "annotate 操作必须提供 outputPath。");
    if (!string.Equals(Path.GetExtension(request.OutputPath), ".docx", StringComparison.OrdinalIgnoreCase))
        return ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "输出文件必须是 .docx。");

    var outputPath = Path.GetFullPath(request.OutputPath);
    var pathComparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
    if (string.Equals(Path.GetFullPath(fullPath), outputPath, pathComparison))
        return ProtocolHostResponse.Error(request.RequestId, EngineProtocol.InvalidRequest, "输出路径必须与原文不同。");
    if (File.Exists(outputPath))
        return ProtocolHostResponse.Error(request.RequestId, EngineProtocol.PatchConflict,
            "输出文件已存在：" + outputPath + "。为避免覆盖已有文件，请换一个文件名。");

    var analysis = Analyze(fullPath);
    WordAnnotationWriter.WriteAnnotatedCopy(fullPath, outputPath, analysis.Diagnostics);
    return ProtocolHostResponse.Annotated(request.RequestId, outputPath, analysis);
}

static List<Patch> SelectPatches(PatchPlan plan, EngineRequest request)
{
    if (string.IsNullOrWhiteSpace(request.RuleId) && string.IsNullOrWhiteSpace(request.Location) &&
        (request.Locations == null || request.Locations.Count == 0))
        return plan.Patches.ToList();

    var locations = request.Locations?.Where(location => !string.IsNullOrWhiteSpace(location))
        .ToHashSet(StringComparer.Ordinal) ?? new HashSet<string>(StringComparer.Ordinal);
    if (locations.Count == 0 && !string.IsNullOrWhiteSpace(request.Location))
        locations.Add(request.Location);

    return plan.Patches.Where(patch =>
        (string.IsNullOrWhiteSpace(request.RuleId) || string.Equals(patch.RuleId, request.RuleId, StringComparison.OrdinalIgnoreCase)) &&
        (locations.Count == 0 || locations.Contains(patch.Location)))
        .ToList();
}

static bool TryResolveInput(EngineRequest request, out string fullPath, out string error)
{
    fullPath = string.Empty;
    error = string.Empty;
    if (string.IsNullOrWhiteSpace(request.InputPath) ||
        !string.Equals(Path.GetExtension(request.InputPath), ".docx", StringComparison.OrdinalIgnoreCase))
    {
        error = "输入文件必须是 .docx。";
        return false;
    }

    fullPath = Path.GetFullPath(request.InputPath);
    if (!File.Exists(fullPath))
    {
        error = "找不到输入文件：" + fullPath;
        return false;
    }

    return true;
}

static void Write(EngineResponse response)
{
    Console.WriteLine(EngineProtocolCodec.SerializeResponse(response));
    Console.Out.Flush();
}

static class ProtocolHostResponse
{
    public static EngineResponse Analysis(string requestId, AnalysisResult result)
    {
        var plan = PatchEngine.PlanPatches(result);
        return new EngineResponse {
            ProtocolVersion = EngineProtocol.Version,
            RequestId = requestId, Ok = true, Result = ToAnalysisTransfer(result),
            // Reuse this exact analysis, rather than launching the analyzer again
            // on the next click. Word still checks a complete live fingerprint.
            Plan = ToPlanTransfer(plan, plan.Patches.ToList(), result, new EngineRequest())
        };
    }

    public static EngineResponse Plan(string requestId, EnginePatchPlan plan) => new()
    {
        ProtocolVersion = EngineProtocol.Version,
        RequestId = requestId,
        Ok = true,
        Plan = plan
    };

    public static EngineResponse Applied(string requestId, ApplyResult result) => new()
    {
        ProtocolVersion = EngineProtocol.Version,
        RequestId = requestId,
        Ok = true,
        Apply = new EngineApplyResult
        {
            OutputPath = result.OutputPath ?? string.Empty,
            Applied = result.AppliedPatches.Select(ToPatchTransfer).ToList(),
            Failures = result.Failures.Select(failure => new EnginePatchFailure
            {
                Patch = ToPatchTransfer(failure.Patch),
                Reason = failure.Reason
            }).ToList(),
            Unresolved = result.UnresolvedPatches.Select(ToPatchTransfer).ToList()
        }
    };

    public static EngineResponse Annotated(string requestId, string outputPath, AnalysisResult analysis) => new()
    {
        ProtocolVersion = EngineProtocol.Version,
        RequestId = requestId,
        Ok = true,
        Result = ToAnalysisTransfer(analysis),
        Annotation = new EngineAnnotationResult
        {
            OutputPath = outputPath,
            CommentCount = analysis.Diagnostics?.Count ?? 0
        }
    };

    public static EngineResponse Error(string requestId, string code, string message) => new()
    {
        ProtocolVersion = EngineProtocol.Version,
        RequestId = requestId,
        Ok = false,
        Error = new EngineError { Code = code, Message = message }
    };

    public static EngineAnalysisResult ToAnalysisTransfer(AnalysisResult result) => new()
    {
        RuleVersion = result.RuleVersion,
        Coverage = result.Coverage,
        UnsupportedScopes = result.UnsupportedScopes?.ToList() ?? new List<string>(),
        Diagnostics = (result.Diagnostics ?? Array.Empty<Diagnostic>()).Select(ToDiagnosticTransfer).ToList()
    };

    /// <summary>
    /// 计划的转移表示：Patches 是本次会执行的部分；Deferred 是留给用户的诊断
    /// （既包括引擎判定不能自动修复的，也包括被 ruleId/location 过滤掉的 auto 项）。
    /// </summary>
    public static EnginePatchPlan ToPlanTransfer(PatchPlan plan, List<Patch> selected, AnalysisResult analysis, EngineRequest request)
    {
        var selectedKeys = selected.Select(patch => patch.RuleId + "|" + patch.Location).ToList();
        var deferred = plan.DeferredDiagnostics.ToList();
        bool filtered = !string.IsNullOrWhiteSpace(request.RuleId) || !string.IsNullOrWhiteSpace(request.Location) ||
            request.Locations is { Count: > 0 };
        if (filtered)
        {
            deferred.AddRange((analysis.Diagnostics ?? Array.Empty<Diagnostic>()).Where(diagnostic =>
                string.Equals(diagnostic.FixLevel, "auto", StringComparison.OrdinalIgnoreCase) &&
                !selectedKeys.Contains(diagnostic.RuleId + "|" + diagnostic.Location)));
        }

        return new EnginePatchPlan
        {
            Patches = selected.Select(ToPatchTransfer).ToList(),
            Deferred = deferred.Select(ToDiagnosticTransfer).ToList(),
            RuleVersion = analysis.RuleVersion,
            Coverage = analysis.Coverage,
            UnsupportedScopes = analysis.UnsupportedScopes?.ToList() ?? new List<string>()
        };
    }

    public static EngineDiagnostic ToDiagnosticTransfer(Diagnostic diagnostic) => new()
    {
        RuleId = diagnostic.RuleId,
        Severity = diagnostic.Severity,
        Location = diagnostic.Location,
        Expected = diagnostic.Expected,
        Actual = diagnostic.Actual,
        FixLevel = diagnostic.FixLevel,
        Message = diagnostic.Message,
        Occurrences = (diagnostic.Occurrences ?? Array.Empty<DiagnosticOccurrence>())
            .Select(occurrence => new EngineDiagnosticOccurrence { Location = occurrence.Location, Actual = occurrence.Actual })
            .ToList()
    };

    public static EnginePatch ToPatchTransfer(Patch patch) => new()
    {
        RuleId = patch.RuleId,
        Location = patch.Location,
        ExpectedBefore = patch.ExpectedBefore,
        DesiredAfter = patch.DesiredAfter,
        Reason = patch.Reason
    };
}
