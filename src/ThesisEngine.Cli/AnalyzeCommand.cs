namespace ThesisEngine.Cli
{
using ThesisEngine.Core;

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

    public int Execute(
        IEnumerable<string> arguments,
        TextWriter output,
        TextWriter error)
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

        if (!TryParse(arguments, out var path))
        {
            WriteUsage(error);
            return 2;
        }

        string extension;
        try
        {
            extension = System.IO.Path.GetExtension(path);
        }
        catch (Exception)
        {
            error.WriteLine("输入路径无效，请检查路径格式。");
            return 1;
        }

        if (!string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
        {
            error.WriteLine($"输入格式不受支持：仅支持 .docx 文件。文件：{path}");
            return 1;
        }

        if (!File.Exists(path))
        {
            error.WriteLine($"找不到输入文件：{path}");
            return 1;
        }

        try
        {
            var result = _analysisService.Analyze(path, CreateDemoProfile());
            if (result is null)
            {
                error.WriteLine("分析失败：分析服务未返回结果。");
                return 1;
            }

            CliReportFormatter.Write(result, output, path);
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
        catch (Exception)
        {
            error.WriteLine("分析失败，请检查文档格式或稍后重试。");
            return 1;
        }
    }

    private static bool TryParse(IReadOnlyList<string> arguments, out string path)
    {
        path = "";
        if (arguments.Count != 3 ||
            !string.Equals(arguments[0], "analyze", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(arguments[1], "--demo", StringComparison.Ordinal) ||
            string.IsNullOrWhiteSpace(arguments[2]))
        {
            return false;
        }

        path = arguments[2];
        return true;
    }

    private static DemoRuleSet CreateDemoProfile() => DemoRuleSet.NkuDraftSample;

    private static void WriteUsage(TextWriter error)
    {
        error.WriteLine("用法：ThesisEngine.Cli analyze --demo <file.docx>");
    }
}
}
