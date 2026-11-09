using System.Globalization;
using System.Text;
using ThesisEngine.Core.Rules;

namespace ThesisEngine.Core.Analysis;

/// <summary>
/// Compares the loaded YAML rules against a document snapshot and a paragraph/section
/// layout probe. Every check is one of 符合 / 不符合 / 未检查; a missing or unresolved
/// value can never become a pass.
/// </summary>
public sealed class ThesisRuleAnalyzer
{
    public RuleAnalysisResult Analyze(DocumentSnapshot snapshot, DocumentLayoutProbe layout, AnalysisRuleSet ruleSet)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(layout);
        ArgumentNullException.ThrowIfNull(ruleSet);

        var mapped = ThesisObjectMap.Create(snapshot, layout);
        var byKind = mapped
            .Where(item => item.Checkable)
            .GroupBy(item => item.ObjectKind, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.OrderBy(item => item.ParagraphIndex).ToArray(), StringComparer.Ordinal);
        var findings = new List<RuleFinding>();
        var uncheckedItems = new List<UncheckedItem>();
        var coverage = new List<RuleCoverageRow>();
        var passed = 0;
        var failed = 0;
        var suppressed = 0;
        var checkedParagraphs = new HashSet<int>();
        var ruleOrder = new Dictionary<string, int>(StringComparer.Ordinal);

        for (var index = 0; index < ruleSet.Rules.Count; index++)
        {
            var rule = ruleSet.Rules[index];
            ruleOrder[rule.RuleId] = index;
            if (!rule.Implemented)
            {
                coverage.Add(new RuleCoverageRow(rule.RuleId, rule.YamlPath, rule.Clause, rule.ObjectKind,
                    rule.Property, rule.ExpectedDisplay, false, rule.UncheckedReason ?? "未实现", 0, 0));
                uncheckedItems.Add(new UncheckedItem(rule.RuleId, rule.YamlPath, rule.Clause, null, null,
                    rule.UncheckedReason ?? "规则未实现"));
                continue;
            }

            var applicable = 0;
            var failures = 0;
            var ruleSuppressed = 0;
            if (string.Equals(rule.ObjectKind, AnalysisObjectKinds.Page, StringComparison.Ordinal))
            {
                foreach (var section in layout.Sections)
                {
                    var outcome = CheckSection(rule, section);
                    if (outcome is null) continue;
                    applicable++;
                    if (outcome.Status == RuleOutcomeStatus.Fail)
                    {
                        failures++;
                        failed++;
                        if (rule.Report)
                            findings.Add(ToSectionFinding(rule, ruleSet, section, outcome));
                        else
                        {
                            ruleSuppressed++;
                            suppressed++;
                        }
                    }
                    else if (outcome.Status == RuleOutcomeStatus.Unchecked)
                    {
                        uncheckedItems.Add(new UncheckedItem(rule.RuleId, rule.YamlPath, rule.Clause,
                            $"section[{section.Index}]", null, outcome.Reason ?? "无法判定"));
                    }
                    else
                    {
                        passed++;
                    }
                }
            }
            else if (byKind.TryGetValue(rule.ObjectKind, out var targets))
            {
                foreach (var target in targets)
                {
                    var outcome = CheckParagraph(rule, target);
                    if (outcome is null) continue;
                    applicable++;
                    checkedParagraphs.Add(target.ParagraphIndex);
                    if (outcome.Status == RuleOutcomeStatus.Fail)
                    {
                        failures++;
                        failed++;
                        if (rule.Report)
                            findings.Add(ToParagraphFinding(rule, ruleSet, target, outcome));
                        else
                        {
                            ruleSuppressed++;
                            suppressed++;
                        }
                    }
                    else if (outcome.Status == RuleOutcomeStatus.Unchecked)
                    {
                        uncheckedItems.Add(new UncheckedItem(rule.RuleId, rule.YamlPath, rule.Clause,
                            ParagraphLocation(target), target.ParagraphIndex, outcome.Reason ?? "无法判定"));
                    }
                    else
                    {
                        passed++;
                    }
                }
            }

            if (ruleSuppressed > 0)
                uncheckedItems.Add(new UncheckedItem(rule.RuleId, rule.YamlPath, rule.Clause, null, null,
                    $"按预警开关未逐条报告（{ruleSuppressed} 处不符合）"));
            coverage.Add(new RuleCoverageRow(rule.RuleId, rule.YamlPath, rule.Clause, rule.ObjectKind,
                rule.Property, rule.ExpectedDisplay, true,
                ruleSuppressed > 0 ? "预警开关关闭：只计数，不逐条列出" : null, applicable, failures));
        }

        AddObjectLevelUnchecked(mapped, uncheckedItems);
        foreach (var unsupportedScope in snapshot.UnsupportedScopes)
            uncheckedItems.Add(new UncheckedItem("reader", string.Empty, null, null, null, unsupportedScope));
        var orderedFindings = findings
            .GroupBy(finding => (finding.RuleId, finding.Expected, finding.Actual))
            .Select(group =>
            {
                var items = group
                    .OrderBy(finding => finding.ParagraphIndex ?? -1)
                    .ThenBy(finding => finding.Location, StringComparer.Ordinal)
                    .ToArray();
                var occurrences = items
                    .SelectMany(finding => finding.Occurrences)
                    .OrderBy(occurrence => occurrence.ParagraphIndex ?? -1)
                    .ThenBy(occurrence => occurrence.Location, StringComparer.Ordinal)
                    .ToArray();
                var first = items[0];
                var message = occurrences.Length > 1
                    ? first.Message + $" 同类共 {occurrences.Length} 处，已合并为一条问题。"
                    : first.Message;
                return first with
                {
                    Location = occurrences[0].Location,
                    ParagraphIndex = occurrences[0].ParagraphIndex,
                    RunIndex = occurrences[0].RunIndex,
                    Excerpt = occurrences[0].Excerpt,
                    Occurrences = occurrences,
                    Message = message
                };
            })
            .OrderBy(finding => finding.ParagraphIndex ?? -1)
            .ThenBy(finding => ruleOrder.GetValueOrDefault(finding.RuleId, int.MaxValue))
            .ThenBy(finding => finding.RuleId, StringComparer.Ordinal)
            .ToArray();
        var orderedUnchecked = uncheckedItems
            .GroupBy(item => (item.RuleId, item.Location, item.Reason))
            .Select(group => group.First())
            .OrderBy(item => item.ParagraphIndex ?? -1)
            .ThenBy(item => ruleOrder.GetValueOrDefault(item.RuleId, int.MaxValue))
            .ThenBy(item => item.RuleId, StringComparer.Ordinal)
            .ToArray();

        return new RuleAnalysisResult(
            snapshot.PartName,
            ruleSet,
            coverage.AsReadOnly(),
            orderedFindings,
            orderedUnchecked,
            snapshot.Paragraphs.Count,
            checkedParagraphs.Count,
            passed,
            failed,
            orderedUnchecked.Length)
        {
            SuppressedFindings = suppressed
        };
    }

    private static CheckOutcome? CheckParagraph(AnalysisRule rule, MappedParagraph target)
    {
        var paragraph = target.Paragraph;
        var layout = target.Layout;
        if (layout is null) return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, "无法解析",
            null, "段落格式探测缺失");

        return rule.Property switch
        {
            RuleProperties.FontEastAsia => CheckEastAsiaFont(rule, paragraph),
            RuleProperties.FontLatin => CheckLatinFont(rule, paragraph),
            RuleProperties.FontSize => CheckFontSize(rule, paragraph),
            RuleProperties.Alignment => CheckAlignment(rule, layout),
            RuleProperties.LineSpacing => CheckLineSpacing(rule, layout, paragraph),
            RuleProperties.SpaceBefore => CheckSpacing(rule, layout.SpaceBeforePt),
            RuleProperties.SpaceAfter => CheckSpacing(rule, layout.SpaceAfterPt),
            RuleProperties.FirstLineIndent => CheckFirstLineIndent(rule, layout, paragraph),
            RuleProperties.Bold => CheckBold(rule, target),
            RuleProperties.KeywordBold => CheckKeywordBold(rule, target),
            _ => null
        };
    }

    private static CheckOutcome? CheckEastAsiaFont(AnalysisRule rule, ParagraphSnapshot paragraph)
    {
        var expected = rule.ExpectedValue;
        if (string.IsNullOrWhiteSpace(expected)) return Unchecked(rule, "规则缺少期望字体");
        var runs = paragraph.Runs
            .Where(run => !string.IsNullOrWhiteSpace(run.Text) && ContainsHan(run.Text))
            .ToArray();
        if (runs.Length == 0) return null;

        var mismatches = new List<string>();
        var unknown = false;
        int? runIndex = null;
        foreach (var run in runs)
        {
            if (run.HasUnresolvedStyle || string.IsNullOrWhiteSpace(run.EastAsiaFont))
            {
                unknown = true;
                runIndex ??= run.RunIndex;
                continue;
            }
            if (!string.Equals(run.EastAsiaFont, expected, StringComparison.Ordinal))
            {
                mismatches.Add(run.EastAsiaFont);
                runIndex ??= run.RunIndex;
            }
        }

        if (mismatches.Count > 0)
            return new CheckOutcome(RuleOutcomeStatus.Fail, expected, JoinDistinct(mismatches), runIndex, null);
        return unknown
            ? new CheckOutcome(RuleOutcomeStatus.Unchecked, expected, "存在无法解析的 eastAsia 槽位", runIndex,
                "eastAsia 字体槽位未解析")
            : new CheckOutcome(RuleOutcomeStatus.Pass, expected, expected, null, null);
    }

    private static CheckOutcome? CheckLatinFont(AnalysisRule rule, ParagraphSnapshot paragraph)
    {
        var expected = rule.ExpectedValue;
        if (string.IsNullOrWhiteSpace(expected)) return Unchecked(rule, "规则缺少期望字体");
        var runs = paragraph.Runs
            .Where(run => !string.IsNullOrWhiteSpace(run.Text) && ContainsLatinLetter(run.Text))
            .ToArray();
        if (runs.Length == 0) return null;

        var mismatches = new List<string>();
        var unknown = false;
        int? runIndex = null;
        foreach (var run in runs)
        {
            if (run.HasUnresolvedStyle || string.IsNullOrWhiteSpace(run.AsciiSlotFont) || string.IsNullOrWhiteSpace(run.HAnsiFont))
            {
                unknown = true;
                runIndex ??= run.RunIndex;
                continue;
            }
            if (!string.Equals(run.AsciiSlotFont, expected, StringComparison.Ordinal) ||
                !string.Equals(run.HAnsiFont, expected, StringComparison.Ordinal))
            {
                mismatches.Add(DescribeLatin(run));
                runIndex ??= run.RunIndex;
            }
        }

        if (mismatches.Count > 0)
            return new CheckOutcome(RuleOutcomeStatus.Fail, expected, JoinDistinct(mismatches), runIndex, null);
        return unknown
            ? new CheckOutcome(RuleOutcomeStatus.Unchecked, expected, "存在无法解析的 ascii/hAnsi 槽位", runIndex,
                "ascii 或 hAnsi 字体槽位未解析")
            : new CheckOutcome(RuleOutcomeStatus.Pass, expected, expected, null, null);
    }

    private static string DescribeLatin(TextRunSnapshot run) =>
        string.Equals(run.AsciiSlotFont, run.HAnsiFont, StringComparison.Ordinal)
            ? run.AsciiSlotFont ?? "(unknown)"
            : $"ascii={run.AsciiSlotFont ?? "(unknown)"}; hAnsi={run.HAnsiFont ?? "(unknown)"}";

    private static CheckOutcome? CheckFontSize(AnalysisRule rule, ParagraphSnapshot paragraph)
    {
        if (!double.TryParse(rule.ExpectedValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var expected))
            return Unchecked(rule, "规则缺少期望字号");
        var runs = paragraph.Runs.Where(run => !string.IsNullOrWhiteSpace(run.Text)).ToArray();
        if (runs.Length == 0) return null;

        var mismatches = new List<string>();
        var unknown = false;
        int? runIndex = null;
        foreach (var run in runs)
        {
            if (run.HasUnresolvedStyle || run.FontSizePt is null)
            {
                unknown = true;
                runIndex ??= run.RunIndex;
                continue;
            }
            if (Math.Abs(run.FontSizePt.Value - expected) > 0.01)
            {
                mismatches.Add(FormatNumber(run.FontSizePt.Value) + "pt");
                runIndex ??= run.RunIndex;
            }
        }

        var expectedDisplay = FormatNumber(expected) + "pt";
        if (mismatches.Count > 0)
            return new CheckOutcome(RuleOutcomeStatus.Fail, expectedDisplay, JoinDistinct(mismatches), runIndex, null);
        return unknown
            ? new CheckOutcome(RuleOutcomeStatus.Unchecked, expectedDisplay, "存在无法解析的字号", runIndex, "字号未解析")
            : new CheckOutcome(RuleOutcomeStatus.Pass, expectedDisplay, expectedDisplay, null, null);
    }

    private static CheckOutcome CheckAlignment(AnalysisRule rule, ParagraphLayout layout)
    {
        var expected = rule.ExpectedValue;
        if (string.IsNullOrWhiteSpace(expected))
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, "规则缺少期望对齐", null,
                "规则缺少期望对齐");
        if (string.Equals(layout.Alignment, expected, StringComparison.Ordinal))
            return new CheckOutcome(RuleOutcomeStatus.Pass, DisplayAlignment(expected), DisplayAlignment(layout.Alignment), null, null);
        return new CheckOutcome(RuleOutcomeStatus.Fail, DisplayAlignment(expected), DisplayAlignment(layout.Alignment), null, null);
    }

    private static CheckOutcome CheckLineSpacing(AnalysisRule rule, ParagraphLayout layout, ParagraphSnapshot paragraph)
    {
        if (layout.HasEquation)
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, "公式段落", null,
                "含数学表达式，按 A1 行距豁免保留人工检查");

        var expected = rule.ExpectedValue;
        var actual = DescribeLineSpacing(layout);
        if (string.IsNullOrWhiteSpace(expected))
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, actual, null, "规则缺少期望行距");

        if (expected.StartsWith("exact:", StringComparison.Ordinal) &&
            double.TryParse(expected[6..], NumberStyles.Float, CultureInfo.InvariantCulture, out var exact))
        {
            if (layout.LineSpacingKind == "exact" && layout.LineSpacingPt is { } actualPt &&
                Math.Abs(actualPt - exact) <= 0.1)
                return new CheckOutcome(RuleOutcomeStatus.Pass, rule.ExpectedDisplay, actual, null, null);
            return new CheckOutcome(RuleOutcomeStatus.Fail, rule.ExpectedDisplay, actual, null, null);
        }

        if (string.Equals(expected, "single", StringComparison.OrdinalIgnoreCase))
            return string.Equals(layout.LineSpacingKind, "single", StringComparison.Ordinal)
                ? new CheckOutcome(RuleOutcomeStatus.Pass, rule.ExpectedDisplay, actual, null, null)
                : new CheckOutcome(RuleOutcomeStatus.Fail, rule.ExpectedDisplay, actual, null, null);

        if (expected.StartsWith("multiple:", StringComparison.Ordinal) &&
            double.TryParse(expected[9..], NumberStyles.Float, CultureInfo.InvariantCulture, out var multiple))
        {
            if (layout.LineSpacingKind == "multiple" && layout.LineSpacingMultiple is { } actualMultiple &&
                Math.Abs(actualMultiple - multiple) <= 0.01)
                return new CheckOutcome(RuleOutcomeStatus.Pass, rule.ExpectedDisplay, actual, null, null);
            return new CheckOutcome(RuleOutcomeStatus.Fail, rule.ExpectedDisplay, actual, null, null);
        }

        return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, actual, null, "未识别的期望行距格式");
    }

    private static string DescribeLineSpacing(ParagraphLayout layout) => layout.LineSpacingKind switch
    {
        "single" => "单倍行距",
        "exact" => $"固定值 {FormatNumber(layout.LineSpacingPt ?? 0)}pt",
        "atLeast" => $"最小值 {FormatNumber(layout.LineSpacingPt ?? 0)}pt",
        "multiple" => $"多倍行距 {FormatNumber(layout.LineSpacingMultiple ?? 0)}",
        _ => layout.LineSpacingKind
    };

    private static CheckOutcome CheckSpacing(AnalysisRule rule, double actual)
    {
        if (!double.TryParse(rule.ExpectedValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var expected))
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, FormatNumber(actual) + "pt", null,
                "规则缺少期望段前/段后值");
        var actualDisplay = FormatNumber(actual) + "pt";
        var expectedDisplay = FormatNumber(expected) + "pt";
        return Math.Abs(actual - expected) <= 0.1
            ? new CheckOutcome(RuleOutcomeStatus.Pass, expectedDisplay, actualDisplay, null, null)
            : new CheckOutcome(RuleOutcomeStatus.Fail, expectedDisplay, actualDisplay, null, null);
    }

    private static CheckOutcome CheckFirstLineIndent(AnalysisRule rule, ParagraphLayout layout, ParagraphSnapshot paragraph)
    {
        if (!double.TryParse(rule.ExpectedValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var expectedChars))
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, "规则缺少期望缩进", null,
                "规则缺少期望缩进");

        if (layout.HangingChars is > 0 || layout.HangingPt is > 0)
            return new CheckOutcome(RuleOutcomeStatus.Fail, $"{FormatNumber(expectedChars)} 字符", "悬挂缩进", null, null);

        double? actualChars = layout.FirstLineChars;
        string actualDisplay;
        if (actualChars is not null)
        {
            actualDisplay = $"{FormatNumber(actualChars.Value)} 字符";
        }
        else if (layout.FirstLinePt is { } firstLinePt)
        {
            var size = paragraph.Runs.FirstOrDefault(run => run.FontSizePt is > 0)?.FontSizePt ?? 12;
            actualChars = firstLinePt / size;
            actualDisplay = $"{FormatNumber(actualChars.Value)} 字符（{FormatNumber(firstLinePt)}pt）";
        }
        else
        {
            actualChars = 0;
            actualDisplay = "0 字符";
        }

        return Math.Abs(actualChars.Value - expectedChars) <= 0.1
            ? new CheckOutcome(RuleOutcomeStatus.Pass, $"{FormatNumber(expectedChars)} 字符", actualDisplay, null, null)
            : new CheckOutcome(RuleOutcomeStatus.Fail, $"{FormatNumber(expectedChars)} 字符", actualDisplay, null, null);
    }

    private static CheckOutcome? CheckBold(AnalysisRule rule, MappedParagraph target)
    {
        if (!bool.TryParse(rule.ExpectedValue, out var expected))
            return Unchecked(rule, "规则缺少期望加粗值");
        var runs = target.Paragraph.Runs.Where(run => !string.IsNullOrWhiteSpace(run.Text)).ToArray();
        if (runs.Length == 0) return null;

        var runLayouts = target.Layout!.Runs.ToDictionary(run => run.RunIndex);
        var known = runs
            .Select(run => (Run: run, Layout: runLayouts.GetValueOrDefault(run.RunIndex)))
            .Where(item => item.Layout is { BoldKnown: true, Bold: not null })
            .Select(item => (item.Run, Bold: item.Layout!.Bold!.Value))
            .ToArray();
        if (known.Length == 0)
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, expected ? "加粗" : "不加粗", "无法解析", null,
                "加粗属性未解析");
        var mismatches = known.Where(item => item.Bold != expected).ToArray();
        if (mismatches.Length > 0)
            return new CheckOutcome(RuleOutcomeStatus.Fail, expected ? "加粗" : "不加粗",
                expected ? "存在未加粗文本" : "存在加粗文本", mismatches[0].Run.RunIndex, null);
        return known.Length < runs.Length
            ? new CheckOutcome(RuleOutcomeStatus.Unchecked, expected ? "加粗" : "不加粗", "部分文本无法解析", null,
                "部分文本的加粗属性未解析")
            : new CheckOutcome(RuleOutcomeStatus.Pass, expected ? "加粗" : "不加粗", expected ? "加粗" : "不加粗", null, null);
    }

    private static CheckOutcome? CheckKeywordBold(AnalysisRule rule, MappedParagraph target)
    {
        if (!bool.TryParse(rule.ExpectedValue, out var expected))
            return Unchecked(rule, "规则缺少期望加粗值");
        var labelRuns = target.Paragraph.Runs
            .Where(run => !string.IsNullOrWhiteSpace(run.Text) && IsKeywordLabel(run.Text))
            .ToArray();
        if (labelRuns.Length == 0) return null;

        var runLayouts = target.Layout!.Runs.ToDictionary(run => run.RunIndex);
        var known = labelRuns
            .Select(run => (Run: run, Layout: runLayouts.GetValueOrDefault(run.RunIndex)))
            .Where(item => item.Layout is { BoldKnown: true, Bold: not null })
            .Select(item => (item.Run, Bold: item.Layout!.Bold!.Value))
            .ToArray();
        if (known.Length == 0)
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, expected ? "关键词标签加粗" : "关键词标签不加粗",
                "无法解析", null, "关键词标签加粗属性未解析");
        if (known.Any(item => item.Bold != expected))
            return new CheckOutcome(RuleOutcomeStatus.Fail, expected ? "关键词标签加粗" : "关键词标签不加粗",
                expected ? "关键词标签未加粗" : "关键词标签被加粗",
                known.First(item => item.Bold != expected).Run.RunIndex, null);
        return new CheckOutcome(RuleOutcomeStatus.Pass, expected ? "关键词标签加粗" : "关键词标签不加粗",
            expected ? "关键词标签加粗" : "关键词标签不加粗", null, null);
    }

    private static bool IsKeywordLabel(string text)
    {
        var normalized = text.TrimStart();
        return normalized.StartsWith("关键词", StringComparison.Ordinal) ||
               normalized.StartsWith("关键字", StringComparison.Ordinal) ||
               normalized.StartsWith("Key Words", StringComparison.OrdinalIgnoreCase) ||
               normalized.StartsWith("Keywords", StringComparison.OrdinalIgnoreCase);
    }

    private static CheckOutcome? CheckSection(AnalysisRule rule, SectionLayout section)
    {
        return rule.Property switch
        {
            RuleProperties.PagePaper => CheckPaper(rule, section),
            RuleProperties.PageMargins => CheckMargins(rule, section),
            _ => null
        };
    }

    private static CheckOutcome CheckPaper(AnalysisRule rule, SectionLayout section)
    {
        var expected = rule.ExpectedValue;
        if (section.WidthCm is null || section.HeightCm is null)
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, expected, "未解析纸张尺寸", null, "节纸张尺寸未解析");
        var actual = $"{section.PageSize}（{FormatNumber(section.WidthCm.Value)}×{FormatNumber(section.HeightCm.Value)}cm）";
        return string.Equals(section.PageSize, expected, StringComparison.OrdinalIgnoreCase)
            ? new CheckOutcome(RuleOutcomeStatus.Pass, expected, actual, null, null)
            : new CheckOutcome(RuleOutcomeStatus.Fail, expected, actual, null, null);
    }

    private static CheckOutcome CheckMargins(AnalysisRule rule, SectionLayout section)
    {
        var expected = rule.ExpectedValue.Split(',', StringSplitOptions.TrimEntries);
        if (expected.Length != 4 ||
            !double.TryParse(expected[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var top) ||
            !double.TryParse(expected[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var bottom) ||
            !double.TryParse(expected[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var left) ||
            !double.TryParse(expected[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var right))
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, "规则格式无法解析", null,
                "规则缺少页边距四值");

        if (section.TopCm is null || section.BottomCm is null || section.LeftCm is null || section.RightCm is null)
            return new CheckOutcome(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, "存在未解析页边距", null,
                "节页边距未全部解析");

        var actual = $"上{FormatNumber(section.TopCm.Value)} / 下{FormatNumber(section.BottomCm.Value)} / 左{FormatNumber(section.LeftCm.Value)} / 右{FormatNumber(section.RightCm.Value)}cm";
        var matches = Math.Abs(section.TopCm.Value - top) <= 0.05 &&
                      Math.Abs(section.BottomCm.Value - bottom) <= 0.05 &&
                      Math.Abs(section.LeftCm.Value - left) <= 0.05 &&
                      Math.Abs(section.RightCm.Value - right) <= 0.05;
        return matches
            ? new CheckOutcome(RuleOutcomeStatus.Pass, rule.ExpectedDisplay, actual, null, null)
            : new CheckOutcome(RuleOutcomeStatus.Fail, rule.ExpectedDisplay, actual, null, null);
    }

    private static RuleFinding ToParagraphFinding(AnalysisRule rule, AnalysisRuleSet ruleSet, MappedParagraph target, CheckOutcome outcome)
    {
        var excerpt = Excerpt(target.Paragraph.Text);
        var location = $"paragraph[{target.ParagraphIndex}]「{excerpt}」";
        var message = $"D2026 {rule.Clause}：{target.ObjectLabel}{Describe(rule.Property)}应为「{outcome.Expected}」，实际为「{outcome.Actual}」。规则仍为 {ruleSet.DisplayStatus}，需人工复核。";
        return new RuleFinding(rule.RuleId, rule.GroupId, rule.YamlPath, rule.Clause, location,
            target.ParagraphIndex, outcome.RunIndex, excerpt, outcome.Expected, outcome.Actual, message,
            Confidence(rule, ruleSet, target.Confidence), rule.Severity, rule.FixLevel)
        {
            Occurrences = new[]
            {
                new RuleFindingOccurrence(location, target.ParagraphIndex, outcome.RunIndex, excerpt, outcome.Actual)
            }
        };
    }

    private static RuleFinding ToSectionFinding(AnalysisRule rule, AnalysisRuleSet ruleSet, SectionLayout section, CheckOutcome outcome)
    {
        var location = $"section[{section.Index}]（paragraph[{section.FirstParagraphIndex}]–[{section.LastParagraphIndex}]）";
        var message = $"D2026 {rule.Clause}：第 {section.Index + 1} 节{Describe(rule.Property)}应为「{outcome.Expected}」，实际为「{outcome.Actual}」。规则仍为 {ruleSet.DisplayStatus}，需人工复核。";
        return new RuleFinding(rule.RuleId, rule.GroupId, rule.YamlPath, rule.Clause, location,
            null, null, string.Empty, outcome.Expected, outcome.Actual, message,
            Confidence(rule, ruleSet, 0.98), rule.Severity, rule.FixLevel)
        {
            Occurrences = new[]
            {
                new RuleFindingOccurrence(location, null, null, string.Empty, outcome.Actual)
            }
        };
    }

    private static string Describe(string property) => property switch
    {
        RuleProperties.FontEastAsia => "中文字体",
        RuleProperties.FontLatin => "西文字体",
        RuleProperties.FontSize => "字号",
        RuleProperties.Alignment => "对齐方式",
        RuleProperties.LineSpacing => "行距",
        RuleProperties.SpaceBefore => "段前距",
        RuleProperties.SpaceAfter => "段后距",
        RuleProperties.FirstLineIndent => "首行缩进",
        RuleProperties.Bold => "加粗",
        RuleProperties.KeywordBold => "关键词标签",
        RuleProperties.PagePaper => "纸张",
        RuleProperties.PageMargins => "页边距",
        _ => property
    };

    private static double Confidence(AnalysisRule rule, AnalysisRuleSet ruleSet, double structureConfidence)
    {
        var value = rule.BaseConfidence * (0.85 + 0.15 * Math.Clamp(structureConfidence, 0, 1));
        if (ruleSet.IsDraft) value = Math.Min(value, 0.75);
        return Math.Round(Math.Clamp(value, 0.05, 0.95), 2);
    }

    private static void AddObjectLevelUnchecked(IReadOnlyList<MappedParagraph> mapped, List<UncheckedItem> items)
    {
        foreach (var group in mapped.Where(item => !item.Checkable)
                     .GroupBy(item => (item.ObjectKind, item.ObjectLabel, item.UncheckReason)))
        {
            var indices = group.Select(item => item.ParagraphIndex).OrderBy(index => index).ToArray();
            var location = indices.Length == 1
                ? $"paragraph[{indices[0]}]"
                : $"paragraph[{indices[0]}–{indices[^1]}]（{indices.Length} 段）";
            items.Add(new UncheckedItem("structure", string.Empty, null, location,
                indices.Length == 1 ? indices[0] : null,
                group.Key.UncheckReason ?? "未检查"));
        }
    }

    private static CheckOutcome Unchecked(AnalysisRule rule, string reason) =>
        new(RuleOutcomeStatus.Unchecked, rule.ExpectedDisplay, "无法判定", null, reason);

    private static string ParagraphLocation(MappedParagraph target) =>
        $"paragraph[{target.ParagraphIndex}]「{Excerpt(target.Paragraph.Text)}」";

    private static string Excerpt(string text)
    {
        var normalized = string.Join(' ', text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return normalized.Length <= 20 ? normalized : normalized[..20] + "…";
    }

    private static string JoinDistinct(IEnumerable<string> values) =>
        string.Join("、", values.Distinct(StringComparer.Ordinal));

    private static string DisplayAlignment(string value) => value switch
    {
        "left" => "左对齐",
        "center" => "居中",
        "right" => "右对齐",
        "justify" => "两端对齐",
        _ => value
    };

    private static string FormatNumber(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);

    private static bool ContainsHan(string text) => text.EnumerateRunes().Any(rune =>
    {
        var value = rune.Value;
        return value is >= 0x3400 and <= 0x4DBF or
               >= 0x4E00 and <= 0x9FFF or
               >= 0x20000 and <= 0x2FA1F;
    });

    private static bool ContainsLatinLetter(string text) => text.EnumerateRunes().Any(rune =>
        Rune.IsLetter(rune) && (rune.Value is >= 0x0041 and <= 0x024F or
                                >= 0x1E00 and <= 0x1EFF or
                                >= 0x2C60 and <= 0x2C7F or
                                >= 0xA720 and <= 0xA7FF or
                                >= 0xAB30 and <= 0xAB6F));

    private sealed record CheckOutcome(
        RuleOutcomeStatus Status,
        string Expected,
        string Actual,
        int? RunIndex,
        string? Reason);
}
