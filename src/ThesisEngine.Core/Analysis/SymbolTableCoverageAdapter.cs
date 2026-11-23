using ThesisEngine.Core.Rules;

namespace ThesisEngine.Core.Analysis;

/// <summary>
/// Bridges the already-implemented R16 recognition in <see cref="BackMatterFormatAnalyzer"/>
/// into the YAML-backed analyze coverage table. The legacy catalog used to mark all seven
/// R16 slots as unimplemented even though the back-matter analyser recognises the
/// symbol-table / CV section and checks its body formatting.
///
/// Design rules (consistent with the rest of the analyzer):
/// - Only findings the source analyser actually reported become failures.
/// - Dimensions the source analyser merges (line spacing + space before/after share one
///   issue code) are mapped to line_spacing only; space_before / space_after stay
///   "未检查" with an explicit reason instead of being silently passed.
/// - The heading-reuse rule is a duplicate of R7 and is deliberately left 未检查.
/// - If the symbol/CV part is absent, every R16 slot stays 未检查 with the source's reason.
/// </summary>
internal static class SymbolTableCoverageAdapter
{
    private const string SymbolRuleIdPrefix = "R16/";

    // BackMatterFormatAnalyzer: CodeStart=10, SymbolOffsets={0,1,2,3,0,0,4}，
    // 数组下标 1..6 = 中文字体/西文字体/字号/对齐/首行缩进/行距与段距。
    // 实际码：B11 中文字体 / B12 西文字体 / B13 字号 / B14 行距与段距合并。
    // （B10 对应下标 0 未使用；对齐与首行缩进在 R16 spec 中未启用，无对应槽。）
    private static readonly Dictionary<string, string[]> CodeToRuleIds = new(StringComparer.Ordinal)
    {
        ["B11"] = new[] { "R16/font.cn" },
        ["B12"] = new[] { "R16/font.western" },
        ["B13"] = new[] { "R16/size_pt" },
        ["B14"] = new[] { "R16/line_spacing" },
    };

    private const string MergedSpacingReason = "行距与段前/段后由同一诊断项合并判定，未单独拆分到本条；不构成符合结论";

    public static RuleAnalysisResult Merge(RuleAnalysisResult result, string docxPath, AnalysisRuleSet ruleSet)
    {
        var r16Rules = ruleSet.Rules
            .Where(r => r.RuleId.StartsWith(SymbolRuleIdPrefix, StringComparison.Ordinal))
            .ToArray();
        if (r16Rules.Length == 0) return result;

        var backMatter = new BackMatterFormatAnalyzer().Analyze(docxPath);
        var symbolIssues = backMatter.Issues
            .Where(i => i.Section == "符号表与个人简历")
            .ToArray();
        var symbolSpan = backMatter.Sections
            .FirstOrDefault(s => s.RuleId == "R16");
        var partRecognized = symbolSpan is not null;

        var findings = new List<RuleFinding>(result.Findings);
        var coverage = new List<RuleCoverageRow>(result.Coverage);
        // 基线把这 7 条记为 Implemented=false，已各生成一条未检查项；先移除，下面按新结论重建。
        var retainedUnchecked = result.UncheckedItems
            .Where(u => !u.RuleId.StartsWith(SymbolRuleIdPrefix, StringComparison.Ordinal))
            .ToArray();
        var removedUnchecked = result.UncheckedItems.Count - retainedUnchecked.Length;
        var uncheckedItems = new List<UncheckedItem>(retainedUnchecked);
        var passed = result.PassedChecks;
        var failed = result.FailedChecks;

        foreach (var rule in r16Rules)
        {
            var index = coverage.FindIndex(c => string.Equals(c.RuleId, rule.RuleId, StringComparison.Ordinal));
            if (index < 0) continue;
            var prior = coverage[index];

            // heading 是 R7 的重复条款，沿用 R7 判定，本条始终未检查。
            if (string.Equals(rule.RuleId, "R16/heading", StringComparison.Ordinal))
            {
                uncheckedItems.Add(new UncheckedItem(rule.RuleId, rule.YamlPath, rule.Clause, null, null,
                    "符号表/个人简历标题复用 R7 章标题规则，见 R7 覆盖行，本条不重复判定"));
                continue;
            }

            // 空间距两条被合并进 B14，只映射到 line_spacing；其余保持未检查。
            if (rule.RuleId is "R16/spacing.before_pt" or "R16/spacing.after_pt")
            {
                uncheckedItems.Add(new UncheckedItem(rule.RuleId, rule.YamlPath, rule.Clause, null, null, MergedSpacingReason));
                continue;
            }

            if (!partRecognized)
            {
                uncheckedItems.Add(new UncheckedItem(rule.RuleId, rule.YamlPath, rule.Clause, null, null,
                    "未识别到符号注释表/个人简历标题段落，正文格式未检查"));
                continue;
            }

            var mapped = symbolIssues
                .Where(i => CodeToRuleIds.TryGetValue(i.Code, out var ids) && ids.Contains(rule.RuleId, StringComparer.Ordinal))
                .ToArray();

            if (mapped.Length == 0)
            {
                // 源分析器对该维度没有报错：只有在其"未检查"列表没有覆盖该维度时才记符合。
                // BackMatterFormatAnalyzer 对无法判定的维度写进 NotChecked；这里保守地按未检查处理。
                var dimensionKey = PropertyDimensionKey(rule.Property);
                var uncertain = backMatter.NotChecked.Any(n =>
                    n.StartsWith("符号表与个人简历", StringComparison.Ordinal) && n.Contains(dimensionKey, StringComparison.Ordinal));
                if (uncertain)
                {
                    uncheckedItems.Add(new UncheckedItem(rule.RuleId, rule.YamlPath, rule.Clause,
                        $"paragraph[{string.Join(",", symbolSpan!.BodyParagraphs)}]", null,
                        "源分析器对该维度列未判定，不据此判合规"));
                }
                else
                {
                    passed++;
                }
                coverage[index] = prior with
                {
                    Implemented = true,
                    Reason = null,
                    Applicable = symbolSpan!.BodyParagraphs.Count,
                    Failures = 0,
                };
                continue;
            }

            foreach (var issue in mapped)
            {
                var location = $"paragraph[{issue.ParagraphIndex}]";
                var message = $"D2026 {rule.Clause}：符号表/个人简历（R16）{Describe(rule.Property)}应为「{issue.Expected}」，实际为「{issue.Actual}」。{issue.Message}";
                findings.Add(new RuleFinding(
                    rule.RuleId, rule.GroupId, rule.YamlPath, rule.Clause, location,
                    issue.ParagraphIndex, null, location, issue.Expected, issue.Actual, message,
                    Math.Min(rule.BaseConfidence, 0.75), rule.Severity, rule.FixLevel)
                {
                    Occurrences = new[]
                    {
                        new RuleFindingOccurrence(location, issue.ParagraphIndex, null, location, issue.Actual)
                    }
                });
                failed++;
            }

            coverage[index] = prior with
            {
                Implemented = true,
                Reason = null,
                Applicable = symbolSpan!.BodyParagraphs.Count,
                Failures = mapped.Length,
            };
        }

        return result with
        {
            Coverage = coverage,
            Findings = findings,
            UncheckedItems = uncheckedItems,
            PassedChecks = passed,
            FailedChecks = failed,
            UncheckedChecks = result.UncheckedChecks - removedUnchecked + uncheckedItems.Count - retainedUnchecked.Length,
        };
    }

    private static string PropertyDimensionKey(string property) => property switch
    {
        RuleProperties.FontEastAsia => "中文字体",
        RuleProperties.FontLatin => "西文字体",
        RuleProperties.FontSize => "字号",
        RuleProperties.LineSpacing => "行距",
        _ => property
    };

    private static string Describe(string property) => property switch
    {
        RuleProperties.FontEastAsia => "中文字体",
        RuleProperties.FontLatin => "西文字体",
        RuleProperties.FontSize => "字号",
        RuleProperties.LineSpacing => "行距",
        RuleProperties.SpaceBefore => "段前距",
        RuleProperties.SpaceAfter => "段后距",
        _ => property
    };
}
