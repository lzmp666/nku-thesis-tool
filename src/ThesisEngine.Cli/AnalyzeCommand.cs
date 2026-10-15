using ThesisEngine.Core;
using System.Xml;

namespace ThesisEngine.Cli;

public sealed class AnalyzeCommand(IDemoAnalysisService analysisService, TextWriter? output = null, TextWriter? error = null)
{
    public int Execute(IEnumerable<string> arguments)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        return Execute(arguments.ToArray(), output ?? Console.Out, error ?? Console.Error);
    }

    public int Execute(string[] arguments) => Execute(arguments, output ?? Console.Out, error ?? Console.Error);

    public int Execute(string[] arguments, TextWriter output) => Execute(arguments, output, error ?? Console.Error);

    public int Execute(IEnumerable<string> arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        return Execute(arguments.ToArray(), output, error);
    }

    public int Execute(string[] arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);

        if (!TryParse(arguments, out var inputPath, out var annotatedOutputPath))
        {
            error.WriteLine("用法：ThesisEngine.Cli analyze --demo <file.docx> [--annotated-output <annotated.docx>]");
            return 2;
        }

        if (!string.Equals(Path.GetExtension(inputPath), ".docx", StringComparison.OrdinalIgnoreCase))
        {
            error.WriteLine($"输入错误：仅支持 .docx 文件：{inputPath}");
            return 1;
        }

        if (!File.Exists(inputPath))
        {
            error.WriteLine($"输入错误：找不到输入文件：{inputPath}");
            return 1;
        }

        if (annotatedOutputPath is not null)
        {
            if (!string.Equals(Path.GetExtension(annotatedOutputPath), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                error.WriteLine($"输出错误：标注副本必须使用 .docx 文件：{annotatedOutputPath}");
                return 1;
            }

            var fullInputPath = Path.GetFullPath(inputPath);
            var fullOutputPath = Path.GetFullPath(annotatedOutputPath);
            var comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            if (string.Equals(fullInputPath, fullOutputPath, comparison))
            {
                error.WriteLine("输出错误：输出路径不能与输入路径相同；为保护原文件，未执行写入。");
                return 1;
            }

            if (File.Exists(fullOutputPath))
            {
                error.WriteLine($"输出错误：目标文件已存在，为避免覆盖，未写入：{annotatedOutputPath}");
                return 1;
            }
        }

        try
        {
            var result = analysisService.Analyze(inputPath, DemoRuleSet.NkuDraftSample);
            var writtenOutputPath = annotatedOutputPath is null
                ? null
                : WordAnnotationWriter.WriteAnnotatedCopy(inputPath, annotatedOutputPath, result.Diagnostics);

            CliReportFormatter.Write(result, output, inputPath, writtenOutputPath);
            return 0;
        }
        catch (UnauthorizedAccessException)
        {
            error.WriteLine("无法读取输入文件，请检查访问权限。");
            return 1;
        }
        catch (IOException)
        {
            error.WriteLine("无法读取文档，请检查文件是否损坏或格式是否正确。");
            return 1;
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidDataException or NotSupportedException or XmlException)
        {
            error.WriteLine($"分析失败：{exception.Message}");
            return 1;
        }
    }

    private static bool TryParse(string[] arguments, out string inputPath, out string? annotatedOutputPath)
    {
        inputPath = string.Empty;
        annotatedOutputPath = null;
        if (arguments.Length < 3 ||
            !string.Equals(arguments[0], "analyze", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(arguments[1], "--demo", StringComparison.Ordinal) ||
            string.IsNullOrWhiteSpace(arguments[2]))
            return false;

        inputPath = arguments[2];
        if (arguments.Length == 3) return true;
        if (arguments.Length != 5 ||
            !string.Equals(arguments[3], "--annotated-output", StringComparison.Ordinal) ||
            string.IsNullOrWhiteSpace(arguments[4]))
            return false;

        annotatedOutputPath = arguments[4];
        return true;
    }
}
