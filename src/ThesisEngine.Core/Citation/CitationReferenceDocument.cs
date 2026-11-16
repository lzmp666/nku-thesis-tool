namespace ThesisEngine.Core.Citation;

public sealed record CitationRunSpan(int Start, int Length, bool IsSuperscript);

public sealed record CitationParagraph(int ParagraphIndex, string Text, IReadOnlyList<CitationRunSpan> Runs);

public sealed record CitationDocument(
    IReadOnlyList<CitationParagraph> Paragraphs,
    int CitationFieldCount);

public sealed record CitationFinding(
    string RuleId,
    int ParagraphIndex,
    string Message,
    string Actual);

public sealed record CitationAnalysisResult(
    IReadOnlyList<CitationFinding> Findings,
    int CitationCount,
    int ReferenceCount,
    int CitationFieldCount,
    IReadOnlyList<string> UnsupportedScopes)
{
    public int Count(string ruleId) => Findings.Count(finding => finding.RuleId == ruleId);
}
