using ThesisEngine.Core;
using System.Xml;

namespace ThesisEngine.Cli;

public sealed class AnalyzeCommand
{
    private readonly IDemoAnalysisService _analysisService;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public AnalyzeCommand(
        IDemoAnalysisService analysisService,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analysisService = analysisService ?? throw new ArgumentNullException(nameof(analysisService));
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

        if (!TryParse(arguments, out var inputPath, out var annotatedOutputPath))
        {
            WriteUsage(error);
            return 2;
        }

        try
        {
            var fullInputPath = Path.GetFullPath(inputPath);
            if (!string.Equals(Path.GetExtension(fullInputPath), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                error.WriteLine($"输入格式不受支持：仅支持 .docx 文件。文件：{inputPath}");
                return 1;
            }

            if (!File.Exists(fullInputPath))
            {
                error.WriteLine($"找不到输入文件：{inputPath}");
                return 1;
            }

            string? fullOutputPath = null;
            if (annotatedOutputPath is not null)
            {
                fullOutputPath = Path.GetFullPath(annotatedOutputPath);
                if (!string.Equals(Path.GetExtension(fullOutputPath), ".docx", StringComparison.OrdinalIgnoreCase))
                {
                    error.WriteLine($"输出格式不受支持：标注副本必须使用 .docx 文件。文件：{annotatedOutputPath}");
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
                    error.WriteLine($"标注输出目标已存在，为避免覆盖，未写入：{annotatedOutputPath}");
                    return 1;
                }

                var outputDirectory = Path.GetDirectoryName(fullOutputPath);
                if (outputDirectory is null || !Directory.Exists(outputDirectory))
                {
                    error.WriteLine($"标注输出目录不存在，请先创建目录：{outputDirectory ?? annotatedOutputPath}");
                    return 1;
                }
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

            CliReportFormatter.Write(result, output, fullInputPath, writtenOutputPath);
            return 0;
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

    private static bool TryParse(IReadOnlyList<string> arguments, out string inputPath, out string? annotatedOutputPath)
    {
        inputPath = string.Empty;
        annotatedOutputPath = null;
        if (arguments.Count < 3 ||
            !string.Equals(arguments[0], "analyze", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(arguments[1], "--demo", StringComparison.Ordinal) ||
            string.IsNullOrWhiteSpace(arguments[2]))
        {
            return false;
        }

        inputPath = arguments[2];
        if (arguments.Count == 3) return true;
        if (arguments.Count != 5 ||
            !string.Equals(arguments[3], "--annotated-output", StringComparison.Ordinal) ||
            string.IsNullOrWhiteSpace(arguments[4]))
        {
            return false;
        }

        annotatedOutputPath = arguments[4];
        return true;
    }

    private static void WriteUsage(TextWriter error) =>
        error.WriteLine("用法：ThesisEngine.Cli analyze --demo <file.docx> [--annotated-output <标注副本.docx>]");
}
