using ThesisEngine.Core.Rules;

namespace ThesisEngine.Core.Analysis;

/// <summary>
/// Read-only composition of the YAML rule loader, the existing formatting reader
/// and the supplementary paragraph/section probe.
/// </summary>
public sealed class RuleAnalysisService : IRuleAnalysisService
{
    private readonly WordFormattingReader _reader;
    private readonly DocxLayoutProbe _layoutProbe;
    private readonly ThesisRuleAnalyzer _analyzer;

    public RuleAnalysisService(
        WordFormattingReader? reader = null,
        DocxLayoutProbe? layoutProbe = null,
        ThesisRuleAnalyzer? analyzer = null)
    {
        _reader = reader ?? new WordFormattingReader(inferUnstyledStructure: true);
        _layoutProbe = layoutProbe ?? new DocxLayoutProbe();
        _analyzer = analyzer ?? new ThesisRuleAnalyzer();
    }

    public RuleAnalysisResult Analyze(string path, string profile = RuleSetLoader.DefaultProfile, string? ruleFile = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentException.ThrowIfNullOrWhiteSpace(profile);
        var ruleSet = RuleSetLoader.Load(profile, ruleFile);
        var snapshot = _reader.Read(path);
        var layout = _layoutProbe.Read(path);
        var result = _analyzer.Analyze(snapshot, layout, ruleSet);
        return result with { InputPath = Path.GetFullPath(path) };
    }
}
