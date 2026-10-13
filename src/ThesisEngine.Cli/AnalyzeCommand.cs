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
        catch (Exception exception)
        {
            error.WriteLine($"Input error: {exception.Message}");
            return 1;
        }

        if (!string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
        {
            error.WriteLine($"Input error: only .docx files are supported: {path}");
            return 1;
        }

        if (!File.Exists(path))
        {
            error.WriteLine($"Input error: file not found: {path}");
            return 1;
        }

        try
        {
            var result = _analysisService.Analyze(path, CreateDemoProfile());
            if (result is null)
            {
                error.WriteLine("Analysis error: the analysis service returned no result.");
                return 1;
            }

            CliReportFormatter.Write(result, output, path);
            return 0;
        }
        catch (Exception exception)
        {
            error.WriteLine($"Analysis error: {exception.Message}");
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
        error.WriteLine("Usage: ThesisEngine.Cli analyze --demo <file.docx>");
    }
}
}
