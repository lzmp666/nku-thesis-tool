using System.Collections.ObjectModel;
namespace ThesisEngine.Core
{

// These small DTOs keep the CLI buildable until the frozen demo contracts are
// supplied by the Core project. They intentionally contain no document logic.
public sealed record DemoRuleSet
{
    public string Version { get; init; } = "demo-0.1";
    public string DraftMarker { get; init; } = "draft/demo";
    public bool IsDraft { get; init; } = true;
    public bool Draft { get => IsDraft; init => IsDraft = value; }
    public string EastAsianFont { get; init; } = "宋体";
    public string EastAsiaFont { get => EastAsianFont; init => EastAsianFont = value; }
    public string WesternFont { get; init; } = "Times New Roman";
    public double FontSizePt { get; init; } = 12;
    public double FontSizeTarget { get => FontSizePt; init => FontSizePt = value; }
    public double TargetFontSizePt { get => FontSizePt; init => FontSizePt = value; }

    public DemoRuleSet() { }

    public DemoRuleSet(
        string version,
        string draftMarker,
        string eastAsianFont,
        string westernFont,
        double fontSizePt)
    {
        Version = version;
        DraftMarker = draftMarker;
        IsDraft = true;
        EastAsianFont = eastAsianFont;
        WesternFont = westernFont;
        FontSizePt = fontSizePt;
    }

    public DemoRuleSet(
        string version,
        bool isDraft,
        string eastAsianFont,
        string westernFont,
        double fontSizePt)
    {
        Version = version;
        DraftMarker = isDraft ? "draft/demo" : "approved";
        IsDraft = isDraft;
        EastAsianFont = eastAsianFont;
        WesternFont = westernFont;
        FontSizePt = fontSizePt;
    }
}

public sealed record TextRunSnapshot
{
    public string Part { get; init; } = "";
    public string Location { get; init; } = "";
    public string Paragraph { get; init; } = "";
    public string Run { get; init; } = "";
    public string ParagraphId { get => Paragraph; init => Paragraph = value; }
    public string RunId { get => Run; init => Run = value; }
    public string? EastAsianFont { get; init; }
    public string? WesternFont { get; init; }
    public double? FontSizePt { get; init; }

    public TextRunSnapshot() { }

    public TextRunSnapshot(
        string part,
        string location,
        string? eastAsianFont,
        string? westernFont,
        double? fontSizePt)
    {
        Part = part;
        Location = location;
        EastAsianFont = eastAsianFont;
        WesternFont = westernFont;
        FontSizePt = fontSizePt;
    }

    public TextRunSnapshot(
        string part,
        string paragraph,
        string run,
        string? eastAsianFont,
        string? westernFont,
        double? fontSizePt)
    {
        Part = part;
        Paragraph = paragraph;
        Run = run;
        Location = $"{paragraph}/{run}";
        EastAsianFont = eastAsianFont;
        WesternFont = westernFont;
        FontSizePt = fontSizePt;
    }
}

public sealed record DocumentSnapshot
{
    public string Path { get; init; } = "";
    public IReadOnlyList<TextRunSnapshot> TextRuns { get; init; } = Array.Empty<TextRunSnapshot>();
    public IReadOnlyList<TextRunSnapshot> Runs { get => TextRuns; init => TextRuns = value; }

    public DocumentSnapshot() { }

    public DocumentSnapshot(string path, IEnumerable<TextRunSnapshot> textRuns)
    {
        Path = path;
        TextRuns = new ReadOnlyCollection<TextRunSnapshot>(textRuns.ToList());
    }
}

public sealed record Diagnostic
{
    public string RuleId { get; init; } = "";
    public string Severity { get; init; } = "";
    public string Location { get; init; } = "";
    public string Expected { get; init; } = "";
    public string Actual { get; init; } = "";
    public string FixLevel { get; init; } = "";
    public string Message { get; init; } = "";

    public Diagnostic() { }

    public Diagnostic(
        string ruleId,
        string severity,
        string location,
        string expected,
        string actual,
        string fixLevel,
        string message)
    {
        RuleId = ruleId;
        Severity = severity;
        Location = location;
        Expected = expected;
        Actual = actual;
        FixLevel = fixLevel;
        Message = message;
    }
}

public sealed record AnalysisResult
{
    public IReadOnlyList<Diagnostic> Diagnostics { get; init; } = Array.Empty<Diagnostic>();
    public string RuleVersion { get; init; } = "";
    public string Coverage { get; init; } = "";
    public IReadOnlyList<string> UnsupportedScopes { get; init; } = Array.Empty<string>();

    public AnalysisResult() { }

    public AnalysisResult(
        IEnumerable<Diagnostic> diagnostics,
        string ruleVersion,
        string coverage,
        IEnumerable<string> unsupportedScopes)
    {
        Diagnostics = new ReadOnlyCollection<Diagnostic>(diagnostics.ToList());
        RuleVersion = ruleVersion;
        Coverage = coverage;
        UnsupportedScopes = new ReadOnlyCollection<string>(unsupportedScopes.ToList());
    }
}

public interface IDemoDocumentReader
{
    DocumentSnapshot Read(string path);
}

public interface IDemoAnalyzer
{
    AnalysisResult Analyze(DocumentSnapshot snapshot, DemoRuleSet profile);
}

public interface IDemoAnalysisService
{
    AnalysisResult Analyze(string path, DemoRuleSet profile);
}
}

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

    private static DemoRuleSet CreateDemoProfile() => new()
    {
        Version = "demo-0.1",
        DraftMarker = "draft/demo",
        IsDraft = true,
        EastAsianFont = "宋体",
        WesternFont = "Times New Roman",
        FontSizePt = 12
    };

    private static void WriteUsage(TextWriter error)
    {
        error.WriteLine("Usage: ThesisEngine.Cli analyze --demo <file.docx>");
    }
}
}
