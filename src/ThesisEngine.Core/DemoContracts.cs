namespace ThesisEngine.Core;

/// <summary>
/// Demo-only 规则集。显式标记为 draft/demo，不能用于正式 NKU 合规检查。
/// 正文段落字体/字号已按 2026 版规范第 4.7 节表 4.6 核对（#43）。
/// </summary>
public sealed record DemoRuleSet(
    string RuleVersion,
    bool IsDraft,
    string EastAsiaFontExpected,
    string AsciiFontExpected,
    double FontSizeExpectedPt)
{
    public static DemoRuleSet NkuDraftSample => new(
        RuleVersion: "demo-0.1.0",
        IsDraft: true,
        EastAsiaFontExpected: "宋体",           // 2026版 4.7表4.6: 中文正文宋体12pt
        AsciiFontExpected: "Times New Roman",  // 2026版 4.7表4.6: 英文正文TNR 12pt
        FontSizeExpectedPt: 12.0);             // 2026版 4.7表4.6: 正文字号12pt
}

/// <summary>
/// 文档快照：只读读取后的中间表示，供诊断器使用。
/// </summary>
public sealed record DocumentSnapshot(
    string PartName,
    IReadOnlyList<ParagraphSnapshot> Paragraphs,
    IReadOnlyList<string> UnsupportedScopes);

public sealed record ParagraphSnapshot(
    int ParagraphIndex,
    IReadOnlyList<TextRunSnapshot> Runs);

public sealed record TextRunSnapshot(
    int RunIndex,
    string? EastAsiaFont,
    string? AsciiFont,
    double? FontSizePt,
    string Text);

/// <summary>
/// 诊断结果条目，对齐 engine-api.md 中的 Diagnostic 字段。
/// </summary>
public sealed record Diagnostic(
    string RuleId,
    string Severity,
    string Location,
    string Expected,
    string Actual,
    string FixLevel,
    string Message);

/// <summary>
/// 分析结果，对齐 engine-api.md 中的 AnalysisResult。
/// </summary>
public sealed record AnalysisResult(
    IReadOnlyList<Diagnostic> Diagnostics,
    string RuleVersion,
    string Coverage,
    IReadOnlyList<string> UnsupportedScopes);

/// <summary>
/// Demo 文档读取器接口。
/// </summary>
public interface IDemoDocumentReader
{
    DocumentSnapshot Read(string path);
}

/// <summary>
/// Demo 规则诊断器接口。
/// </summary>
public interface IDemoAnalyzer
{
    AnalysisResult Analyze(DocumentSnapshot snapshot, DemoRuleSet profile);
}

/// <summary>
/// Demo 分析服务（读取 + 诊断的组合入口）。
/// </summary>
public interface IDemoAnalysisService
{
    AnalysisResult Analyze(string path, DemoRuleSet profile);
}
