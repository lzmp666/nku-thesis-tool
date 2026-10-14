namespace ThesisEngine.Core;

/// <summary>
/// Connects the demo-only document reader and analyzer into the CLI's real analysis path.
/// This service is read-only and does not modify the input document.
/// </summary>
public sealed class DemoAnalysisService : IDemoAnalysisService
{
    private readonly IDemoDocumentReader _reader;
    private readonly IDemoAnalyzer _analyzer;

    public DemoAnalysisService(IDemoDocumentReader reader, IDemoAnalyzer analyzer)
    {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        _analyzer = analyzer ?? throw new ArgumentNullException(nameof(analyzer));
    }

    public AnalysisResult Analyze(string path, DemoRuleSet profile)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(profile);

        var snapshot = _reader.Read(path);
        return _analyzer.Analyze(snapshot, profile);
    }
}
