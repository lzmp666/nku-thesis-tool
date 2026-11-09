namespace ThesisEngine.Core.Rules;

public enum RuleOutcomeStatus
{
    Pass,
    Fail,
    Unchecked
}

/// <summary>
/// One YAML-backed rule slot. A rule may still be unimplemented (for example a
/// page header rule when headers are not read); it is then reported as 未检查,
/// never silently omitted.
/// </summary>
public sealed record AnalysisRule(
    string RuleId,
    string GroupId,
    string YamlPath,
    string Clause,
    string SourceId,
    string ObjectKind,
    string Property,
    string ExpectedDisplay,
    string ExpectedValue,
    string Severity,
    string FixLevel,
    bool Implemented,
    string? UncheckedReason,
    double BaseConfidence,
    string? MatrixId = null)
{
    /// <summary>
    /// Report-level switch. A disabled rule is still checked for coverage, but its
    /// individual locations are not listed; the count remains visible.
    /// </summary>
    public bool Report { get; init; } = true;
}

public sealed record AnalysisRuleSet(
    string ProfileId,
    string RuleVersion,
    string Status,
    string SourcePath,
    IReadOnlyList<AnalysisRule> Rules,
    IReadOnlyList<string> Warnings)
{
    public bool IsDraft => !string.Equals(Status, "approved", StringComparison.OrdinalIgnoreCase);

    public string DisplayStatus => IsDraft ? "draft（未核验）" : "approved";
}

public sealed record RuleFindingOccurrence(
    string Location,
    int? ParagraphIndex,
    int? RunIndex,
    string Excerpt,
    string Actual);

public sealed record RuleFinding(
    string RuleId,
    string GroupId,
    string YamlPath,
    string Clause,
    string Location,
    int? ParagraphIndex,
    int? RunIndex,
    string Excerpt,
    string Expected,
    string Actual,
    string Message,
    double Confidence,
    string Severity,
    string FixLevel)
{
    /// <summary>
    /// Same rule, expected and actual value: the issue is reported once and every
    /// concrete paragraph is retained for navigation and annotation.
    /// </summary>
    public IReadOnlyList<RuleFindingOccurrence> Occurrences { get; init; } = Array.Empty<RuleFindingOccurrence>();
}

public sealed record RuleCoverageRow(
    string RuleId,
    string YamlPath,
    string Clause,
    string ObjectKind,
    string Property,
    string Expected,
    bool Implemented,
    string? Reason,
    int Applicable,
    int Failures)
{
    public string StatusLabel => Implemented ? "已实现" : "未检查";
}

public sealed record UncheckedItem(
    string RuleId,
    string YamlPath,
    string? Clause,
    string? Location,
    int? ParagraphIndex,
    string Reason);

public sealed record RuleAnalysisResult(
    string InputPath,
    AnalysisRuleSet RuleSet,
    IReadOnlyList<RuleCoverageRow> Coverage,
    IReadOnlyList<RuleFinding> Findings,
    IReadOnlyList<UncheckedItem> UncheckedItems,
    int ParagraphCount,
    int CheckedParagraphCount,
    int PassedChecks,
    int FailedChecks,
    int UncheckedChecks)
{
    public int SuppressedFindings { get; init; }

    /// <summary>
    /// Maps paragraph-level findings to the annotation writer's run locations.
    /// Section-level findings have no run anchor and are intentionally excluded.
    /// </summary>
    public IReadOnlyList<Diagnostic> BuildDiagnostics() =>
        Findings
            .SelectMany(finding => (finding.Occurrences.Count > 0
                    ? finding.Occurrences
                    : new[]
                    {
                        new RuleFindingOccurrence(finding.Location, finding.ParagraphIndex, finding.RunIndex,
                            finding.Excerpt, finding.Actual)
                    })
                .Select(occurrence => (Finding: finding, Occurrence: occurrence)))
            .Where(item => item.Occurrence.ParagraphIndex is not null && item.Occurrence.RunIndex is not null)
            .Select(item => new Diagnostic(
                item.Finding.RuleId,
                item.Finding.Severity,
                $"word/document.xml#paragraph[{item.Occurrence.ParagraphIndex}]/run[{item.Occurrence.RunIndex}]",
                item.Finding.Expected,
                item.Occurrence.Actual,
                item.Finding.FixLevel,
                item.Finding.Message))
            .ToArray();
}

public interface IRuleAnalysisService
{
    RuleAnalysisResult Analyze(string path, string profile, string? ruleFile);
}
