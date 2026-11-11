using ThesisEngine.Core;
using ThesisEngine.Core.Analysis;
using ThesisEngine.Core.Document;
using ThesisEngine.Core.Rules;
using System.Xml;

namespace ThesisEngine.Cli;

public sealed class AnalyzeCommand
{
    /// <summary>
    /// #112：文档级报告并入 analyze。默认 all；paragraph 保持老的段落级行为不变。
    /// </summary>
    private enum Scope { All, Paragraph, Document }

    private readonly IDemoAnalysisService? _analysisService;
    private readonly IRuleAnalysisService _ruleService;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public AnalyzeCommand(
        IDemoAnalysisService analysisService,
        TextWriter? output = null,
        TextWriter? error = null)
        : this(analysisService, new RuleAnalysisService(), output, error)
    {
    }

    public AnalyzeCommand(
        IRuleAnalysisService ruleService,
        TextWriter? output = null,
        TextWriter? error = null)
        : this(null, ruleService, output, error)
    {
    }

    private AnalyzeCommand(
        IDemoAnalysisService? analysisService,
        IRuleAnalysisService ruleService,
        TextWriter? output,
        TextWriter? error)
    {
        _analysisService = analysisService;
        _ruleService = ruleService ?? throw new ArgumentNullException(nameof(ruleService));
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(IEnumerable<string> arguments)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        return Execute(arguments.ToArray(), _output, _error);
    }

    public int Execute(IEnumerable<string> arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        return Execute(arguments.ToArray(), output, error);
    }

    public int Execute(string[] arguments)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        return Execute(arguments, _output, _error);
    }

    public int Execute(string[] arguments, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        return Execute(arguments, output, _error);
    }

    public int Execute(string[] arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);

        if (!TryParse(arguments, out var options))
        {
            WriteUsage(error);
            return 2;
        }
        if (options.Help)
        {
            WriteUsage(output);
            return 0;
        }

        try
        {
            var fullInputPath = Path.GetFullPath(options.InputPath);
            if (!string.Equals(Path.GetExtension(fullInputPath), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                error.WriteLine($"输入格式不受支持：仅支持 .docx 文件。文件：{options.InputPath}");
                return 1;
            }
            if (!File.Exists(fullInputPath))
            {
                error.WriteLine($"找不到输入文件：{options.InputPath}");
                return 1;
            }

            string? fullOutputPath = null;
            if (options.AnnotatedOutputPath is not null)
            {
                fullOutputPath = Path.GetFullPath(options.AnnotatedOutputPath);
                if (!string.Equals(Path.GetExtension(fullOutputPath), ".docx", StringComparison.OrdinalIgnoreCase))
                {
                    error.WriteLine($"输出格式不受支持：标注副本必须使用 .docx 文件。文件：{options.AnnotatedOutputPath}");
                    return 1;
                }
                var comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
                if (string.Equals(fullInputPath, fullOutputPath, comparison))
                {
                    error.WriteLine("输出路径不能与输入路径相同；为保护原文件，未执行写入。");
                    return 1;
                }
                if (File.Exists(fullOutputPath) || Directory.Exists(fullOutputPath))
                {
                    error.WriteLine($"标注输出目标已存在，为避免覆盖，未写入：{options.AnnotatedOutputPath}");
                    return 1;
                }
                var outputDirectory = Path.GetDirectoryName(fullOutputPath);
                if (outputDirectory is null || !Directory.Exists(outputDirectory))
                {
                    error.WriteLine($"标注输出目录不存在，请先创建目录：{outputDirectory ?? options.AnnotatedOutputPath}");
                    return 1;
                }
            }

            if (options.Scope == Scope.Document)
            {
                if (options.AnnotatedOutputPath is not null)
                {
                    error.WriteLine("--scope document 只跑文档级，没有段落级诊断可标注；" +
                        "要写标注副本请加 --scope paragraph。");
                    return 2;
                }
                return RunDocumentScope(fullInputPath, output, error);
            }

            if (options.Demo)
            {
                var demoCode = ExecuteDemo(fullInputPath, fullOutputPath, options, output, error);
                if (demoCode != 0)
                    return demoCode;
                return AppendDocumentScope(options, fullInputPath, output, error);
            }

            var result = _ruleService.Analyze(fullInputPath,
                options.Profile ?? RuleSetLoader.DefaultProfile,
                options.RuleFile);
            string? writtenOutputPath = null;
            if (fullOutputPath is not null)
            {
                var diagnostics = result.BuildDiagnostics();
                writtenOutputPath = WordAnnotationWriter.WriteAnnotatedCopy(fullInputPath, fullOutputPath, diagnostics);
            }

            if (options.Json)
                RuleAnalysisReportWriter.WriteJson(result, output);
            else
                RuleAnalysisReportWriter.WriteText(result, output, writtenOutputPath);

            return AppendDocumentScope(options, fullInputPath, output, error);
        }
        catch (UnauthorizedAccessException)
        {
            error.WriteLine("分析失败：无法访问输入或输出文件，请检查文件权限。");
            return 1;
        }
        catch (IOException)
        {
            error.WriteLine("无法读取文档或写入标注副本，请检查文件是否损坏、路径是否有效或目标是否已存在。");
            return 1;
        }
        catch (FormatException exception)
        {
            error.WriteLine($"规则文件格式无效：{exception.Message}");
            return 1;
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidDataException or NotSupportedException or XmlException)
        {
            error.WriteLine($"分析失败：文档、路径或输出操作无效。{exception.Message}");
            return 1;
        }
        catch (Exception)
        {
            error.WriteLine("分析失败，请检查文档格式或稍后重试。");
            return 1;
        }
    }

    /// <summary>单独跑文档级，或者作为 all 的第二段附在段落级报告后面。</summary>
    private static int RunDocumentScope(string path, TextWriter output, TextWriter error) =>
        new DocumentAnalyzeCommand(output: output, error: error).Execute(new[] { path });

    /// <summary>
    /// --scope all 时，在段落级/规则报告之后附一段文档级检查。
    /// paragraph 直接跳过；json 也跳过——`--json` 的契约是 stdout 只有一个 JSON 文档，
    /// 往里混文档级文本或往 stderr 写提示都会破坏机器可读性，改用 usage 说明。
    /// </summary>
    private static int AppendDocumentScope(
        AnalyzeOptions options, string fullInputPath, TextWriter output, TextWriter error)
    {
        if (options.Scope == Scope.Paragraph || options.Json)
            return 0;

        output.WriteLine();
        output.WriteLine("=== 以下为文档级版式检查（analyze 默认附带，--scope paragraph 可关掉）===");
        var documentCode = RunDocumentScope(fullInputPath, output, error);
        if (documentCode != 0)
            output.WriteLine($"（文档级检查未完成，退出码 {documentCode}；上面的段落级报告不受影响）");
        return 0;
    }

    private int ExecuteDemo(string fullInputPath, string? fullOutputPath, AnalyzeOptions options,
        TextWriter output, TextWriter error)
    {
        if (_analysisService is null)
        {
            error.WriteLine("分析失败：当前命令未配置演示分析服务。");
            return 1;
        }
        var result = _analysisService.Analyze(fullInputPath, DemoRuleSet.NkuDraftSample);
        if (result is null)
        {
            error.WriteLine("分析失败：分析服务未返回结果。");
            return 1;
        }
        var writtenOutputPath = fullOutputPath is null
            ? null
            : WordAnnotationWriter.WriteAnnotatedCopy(fullInputPath, fullOutputPath, result.Diagnostics);
        if (options.Json)
        {
            error.WriteLine("用法错误：--json 目前只用于 YAML 规则分析（不带 --demo）。");
            return 2;
        }
        CliReportFormatter.Write(result, output, fullInputPath, writtenOutputPath);
        return 0;
    }

    private static bool TryParse(IReadOnlyList<string> arguments, out AnalyzeOptions options)
    {
        options = default!;
        if (arguments.Count < 2 ||
            !string.Equals(arguments[0], "analyze", StringComparison.OrdinalIgnoreCase))
            return false;

        string? inputPath = null;
        string? annotatedOutputPath = null;
        string? profile = null;
        string? ruleFile = null;
        var scope = Scope.All;
        var demo = false;
        var json = false;
        var help = false;

        for (var index = 1; index < arguments.Count; index++)
        {
            var argument = arguments[index];
            switch (argument)
            {
                case "--demo":
                    demo = true;
                    break;
                case "--json":
                    json = true;
                    break;
                case "--help":
                case "-h":
                    help = true;
                    break;
                case "--profile":
                    if (++index >= arguments.Count || string.IsNullOrWhiteSpace(arguments[index])) return false;
                    profile = arguments[index];
                    break;
                case "--rule-file":
                    if (++index >= arguments.Count || string.IsNullOrWhiteSpace(arguments[index])) return false;
                    ruleFile = arguments[index];
                    break;
                case "--annotated-output":
                    if (++index >= arguments.Count || string.IsNullOrWhiteSpace(arguments[index])) return false;
                    annotatedOutputPath = arguments[index];
                    break;
                case "--scope":
                    if (++index >= arguments.Count) return false;
                    switch (arguments[index])
                    {
                        case "all": scope = Scope.All; break;
                        case "paragraph": scope = Scope.Paragraph; break;
                        case "document": scope = Scope.Document; break;
                        default: return false;
                    }
                    break;
                default:
                    if (argument.StartsWith('-') || !string.IsNullOrWhiteSpace(inputPath)) return false;
                    inputPath = argument;
                    break;
            }
        }

        if (help)
        {
            options = new AnalyzeOptions(string.Empty, null, false, null, null, false, true, Scope.All);
            return true;
        }
        if (string.IsNullOrWhiteSpace(inputPath)) return false;
        if (demo && (profile is not null || ruleFile is not null || json)) return false;
        if (profile is not null && ruleFile is not null) return false;
        options = new AnalyzeOptions(inputPath, annotatedOutputPath, demo, profile, ruleFile, json, false, scope);
        return true;
    }

    private static void WriteUsage(TextWriter writer)
    {
        writer.WriteLine("用法：ThesisEngine.Cli analyze --demo <file.docx> " +
            "[--scope all|paragraph|document] [--annotated-output <标注副本.docx>]");
        writer.WriteLine("      ThesisEngine.Cli analyze <file.docx> [--profile nku-graduate] " +
            "[--rule-file <规则.yaml>] [--scope all|paragraph|document] [--json] " +
            "[--annotated-output <标注副本.docx>]");
        writer.WriteLine("      --scope 默认 all（段落级 + 文档级）；document 只跑文档级；" +
            "--json 只输出规则分析结果，不附文档级。");
    }

    private sealed record AnalyzeOptions(
        string InputPath,
        string? AnnotatedOutputPath,
        bool Demo,
        string? Profile,
        string? RuleFile,
        bool Json,
        bool Help,
        Scope Scope);
}
