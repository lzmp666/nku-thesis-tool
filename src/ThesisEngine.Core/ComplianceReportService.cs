using System.IO.Compression;
using ThesisEngine.Core.Citation;
using ThesisEngine.Core.Document;
using ThesisEngine.Core.Formula;

namespace ThesisEngine.Core;

/// <summary>一条归一化的问题（来自任一分析器）。Code 保留原分析器编号，Analyzer 标明出处。</summary>
public sealed record ComplianceIssue(
    string Analyzer,
    string Code,
    string Title,
    int? ParagraphIndex,
    string Detail,
    bool IsConvention);

/// <summary>某个分析器的执行结果（成功/失败、问题数、未判定数）。</summary>
public sealed record ComplianceSection(
    string Analyzer,
    string Name,
    string? Error,
    int IssueCount,
    int ConventionCount,
    int NotCheckedCount)
{
    public bool Succeeded => Error is null;
}

/// <summary>
/// 统一合规报告：把页面/结构、样式、图表、公式、脚注、引文、摘要、尾注、正文段落、标题层级、目录十一个分析器的结果
/// 汇成一份。严格区分「规范问题」与「排版惯例」与「未判定」，并把单个分析器的失败
/// 记录为 Error 而不是让整份报告崩掉。
/// </summary>
public sealed record ComplianceReport(
    string InputPath,
    IReadOnlyList<ComplianceSection> Sections,
    IReadOnlyList<ComplianceIssue> Issues,
    IReadOnlyList<ComplianceIssue> Conventions,
    IReadOnlyList<string> NotChecked)
{
    public int TotalIssues => Issues.Count;
    public int TotalConventions => Conventions.Count;
    public int TotalNotChecked => NotChecked.Count;
    public int FailedSections => Sections.Count(s => !s.Succeeded);
    public bool AllSucceeded => FailedSections == 0;
}

/// <summary>
/// 汇总全部「只查不改」分析器，输出一份统一报告。
/// 只做聚合与归一化，不新增任何判据、不改任何分析器的结论。
/// </summary>
public sealed class ComplianceReportService
{
    public ComplianceReport Analyze(string docxPath)
    {
        var sections = new List<ComplianceSection>();
        var issues = new List<ComplianceIssue>();
        var conventions = new List<ComplianceIssue>();
        var notChecked = new List<string>();

        void Run(string analyzer, string name, Action collect)
        {
            var issueBefore = issues.Count;
            var conventionBefore = conventions.Count;
            var uncheckedBefore = notChecked.Count;
            string? error = null;
            try
            {
                collect();
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException
                or UnauthorizedAccessException or ArgumentException or System.Xml.XmlException
                or System.IO.InvalidDataException)
            {
                error = ex.Message;
            }
            sections.Add(new ComplianceSection(
                analyzer, name, error,
                issues.Count - issueBefore,
                conventions.Count - conventionBefore,
                notChecked.Count - uncheckedBefore));
        }

        Run("DocumentAnalyzer", "页面与章节结构", () =>
        {
            var result = new DocumentAnalyzer().Analyze(docxPath);
            foreach (var issue in result.Issues)
            {
                var item = new ComplianceIssue(
                    "DocumentAnalyzer", issue.Code, issue.Title, null,
                    $"第 {issue.SectionIndex} 节：期望「{issue.Expected}」，实际「{issue.Current}」；依据 {issue.Basis}",
                    !issue.HasClause);
                (item.IsConvention ? conventions : issues).Add(item);
            }
            notChecked.AddRange(result.UncheckedItems.Select(x => $"页面与章节结构：{x}"));
        });

        Run("StyleAnalyzer", "样式使用诊断", () =>
        {
            var result = new StyleAnalyzer().Analyze(docxPath);
            // 样式分析是诊断信息（未套样式/直接格式覆盖），不是规范违背，不进 Issues；
            // 只把无法归类的无效样式段落作为提示列进未判定。
            foreach (var p in result.InvalidStyleParagraphs)
                notChecked.Add($"样式使用诊断：样式「{p.StyleName}」无法归类，段落「{p.Preview}」，请人工确认");
        });

        Run("FigureTableAnalyzer", "图表排版", () =>
        {
            var result = new FigureTableAnalyzer().Analyze(docxPath);
            foreach (var message in result.PageBreakIssues)
                issues.Add(new ComplianceIssue("FigureTableAnalyzer", "FIG", "图片/表格跨页或排版", null, message, false));
            foreach (var message in result.TableIssues)
                issues.Add(new ComplianceIssue("FigureTableAnalyzer", "TBL", "表格排版", null, message, false));
        });

        Run("FormulaAnalyzer", "公式排版", () =>
        {
            var result = new FormulaAnalyzer().Analyze(docxPath);
            foreach (var issue in result.Issues)
            {
                var location = issue.ParagraphIndex > 0 ? $"第 {issue.ParagraphIndex} 段：" : "";
                var item = new ComplianceIssue(
                    "FormulaAnalyzer", issue.Code, issue.Title, issue.ParagraphIndex,
                    $"{location}期望「{issue.Expected}」，实际「{issue.Current}」；依据 {issue.Basis}",
                    !issue.HasClause);
                (item.IsConvention ? conventions : issues).Add(item);
            }
        });

        Run("FootnoteAnalyzer", "脚注与注释体例", () =>
        {
            var result = new FootnoteAnalyzer().Analyze(docxPath);
            foreach (var issue in result.Issues)
            {
                var detail = issue.ParagraphIndex.HasValue ? $"第 {issue.ParagraphIndex} 段：{issue.Message}" : issue.Message;
                // Code 以 Undetermined 结尾的条目本质是「判不出」，按项目口径归入未判定，
                // 不能混进 Issues 冒充「查出的问题」。
                if (issue.Code.EndsWith("Undetermined", StringComparison.Ordinal))
                    notChecked.Add($"脚注与注释体例：{detail}");
                else
                    issues.Add(new ComplianceIssue(
                        "FootnoteAnalyzer", issue.Code, "注释体例", issue.ParagraphIndex, detail, false));
            }
            notChecked.AddRange(result.NotChecked.Select(x => $"脚注与注释体例：{x}"));
            notChecked.AddRange(result.MissingParts.Select(x => $"脚注与注释体例：缺少部件 {x}"));
        });

        Run("CitationConsistencyAnalyzer", "参考文献引用一致性", () =>
        {
            var result = new CitationConsistencyAnalyzer().Analyze(docxPath);
            foreach (var finding in result.Findings)
                issues.Add(new ComplianceIssue(
                    "CitationConsistencyAnalyzer", finding.RuleId, "引用一致性", finding.ParagraphIndex,
                    $"第 {finding.ParagraphIndex} 段：{finding.Message}（实际「{finding.Actual}」）", false));
            notChecked.AddRange(result.UnsupportedScopes.Select(x => $"参考文献引用一致性：{x}"));
        });

        Run("AbstractFormatAnalyzer", "中英文摘要格式", () =>
        {
            var result = new AbstractFormatAnalyzer().Analyze(docxPath);
            foreach (var issue in result.Issues)
                issues.Add(new ComplianceIssue(
                    "AbstractFormatAnalyzer", issue.Code, issue.Title, issue.ParagraphIndex,
                    $"第 {issue.ParagraphIndex} 段：期望「{issue.Expected}」，实际「{issue.Actual}」；{issue.Message}",
                    false));
            notChecked.AddRange(result.NotChecked.Select(x => $"中英文摘要格式：{x}"));
            notChecked.AddRange(result.MissingParts.Select(x => $"中英文摘要格式：缺少部件 {x}"));
        });

        Run("BackMatterFormatAnalyzer", "尾注部件格式", () =>
        {
            var result = new BackMatterFormatAnalyzer().Analyze(docxPath);
            foreach (var issue in result.Issues)
                issues.Add(new ComplianceIssue(
                    "BackMatterFormatAnalyzer", issue.Code, issue.Title, issue.ParagraphIndex,
                    $"{issue.Section} 第 {issue.ParagraphIndex} 段：期望「{issue.Expected}」，实际「{issue.Actual}」；{issue.Message}",
                    false));
            notChecked.AddRange(result.NotChecked.Select(x => $"尾注部件格式：{x}"));
            notChecked.AddRange(result.MissingParts.Select(x => $"尾注部件格式：缺少部件 {x}"));
        });

        Run("BodyTextFormatAnalyzer", "正文段落格式", () =>
        {
            var result = new BodyTextFormatAnalyzer().Analyze(docxPath);
            foreach (var issue in result.Issues)
                issues.Add(new ComplianceIssue(
                    "BodyTextFormatAnalyzer", issue.Code, issue.Title, issue.ParagraphIndex,
                    $"第 {issue.ParagraphIndex} 段：期望「{issue.Expected}」，实际「{issue.Actual}」；{issue.Message}",
                    false));
            notChecked.AddRange(result.NotChecked.Select(x => $"正文段落格式：{x}"));
        });

        Run("HeadingFormatAnalyzer", "标题层级格式", () =>
        {
            var result = new HeadingFormatAnalyzer().Analyze(docxPath);
            foreach (var issue in result.Issues)
                issues.Add(new ComplianceIssue(
                    "HeadingFormatAnalyzer", issue.Code, $"{issue.Level}{issue.Dimension}", issue.ParagraphIndex,
                    $"第 {issue.ParagraphIndex} 段「{issue.HeadingText}」{issue.Dimension}：期望「{issue.Expected}」，实际「{issue.Actual}」；{issue.Message}",
                    false));
            notChecked.AddRange(result.NotChecked.Select(x => $"标题层级格式：{x}"));
            notChecked.AddRange(result.MissingParts.Select(x => $"标题层级格式：缺少部件 {x}"));
        });

        Run("TocFormatAnalyzer", "目录格式", () =>
        {
            var result = new TocFormatAnalyzer().Analyze(docxPath);
            foreach (var issue in result.Issues)
                issues.Add(new ComplianceIssue(
                    "TocFormatAnalyzer", issue.Code, $"{issue.Part}{issue.Dimension}", issue.ParagraphIndex,
                    $"第 {issue.ParagraphIndex} 段「{issue.Text}」{issue.Dimension}：期望「{issue.Expected}」，实际「{issue.Actual}」；{issue.Message}",
                    false));
            notChecked.AddRange(result.NotChecked.Select(x => $"目录格式：{x}"));
            notChecked.AddRange(result.MissingParts.Select(x => $"目录格式：缺少部件 {x}"));
        });

        return new ComplianceReport(docxPath, sections, issues, conventions, notChecked);
    }
}
