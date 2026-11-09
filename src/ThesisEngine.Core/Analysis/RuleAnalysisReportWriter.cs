using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using ThesisEngine.Core.Rules;

namespace ThesisEngine.Core.Analysis;

public static class RuleAnalysisReportWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public static string ToText(RuleAnalysisResult result, string? annotatedOutputPath = null)
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        WriteText(result, writer, annotatedOutputPath);
        return writer.ToString();
    }

    public static void WriteText(RuleAnalysisResult result, TextWriter writer, string? annotatedOutputPath = null)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(writer);

        var implemented = result.Coverage.Count(row => row.Implemented);
        var uncheckedRules = result.Coverage.Count(row => !row.Implemented);
        writer.WriteLine("=== 论文格式检查报告（规则分析）===");
        writer.WriteLine($"规则集：{result.RuleSet.ProfileId} / {result.RuleSet.RuleVersion} / 状态：{result.RuleSet.DisplayStatus}");
        writer.WriteLine($"规则文件：{result.RuleSet.SourcePath}");
        writer.WriteLine("核验说明：正文 R11 字体/字号来自 S2026；其余组仍为 S1（2025版）草案值，待 2026 全文逐条复核。");
        writer.WriteLine($"输入文件：{result.InputPath}");
        writer.WriteLine("状态：分析完成");
        writer.WriteLine($"段落总数：{result.ParagraphCount.ToString(CultureInfo.InvariantCulture)}；已评估段落：{result.CheckedParagraphCount.ToString(CultureInfo.InvariantCulture)}");
        writer.WriteLine($"检查结果：符合 {result.PassedChecks.ToString(CultureInfo.InvariantCulture)} 项 / 不符合 {result.FailedChecks.ToString(CultureInfo.InvariantCulture)} 项 / 未检查 {result.UncheckedChecks.ToString(CultureInfo.InvariantCulture)} 项");
        writer.WriteLine($"问题数量：{result.Findings.Count.ToString(CultureInfo.InvariantCulture)}（合并后）");
        if (result.SuppressedFindings > 0)
            writer.WriteLine($"预警开关静默：{result.SuppressedFindings.ToString(CultureInfo.InvariantCulture)} 处不符合（只计数，不逐条列出）");
        writer.WriteLine($"规则覆盖：已实现 {implemented.ToString(CultureInfo.InvariantCulture)} 条 / 未检查 {uncheckedRules.ToString(CultureInfo.InvariantCulture)} 条");
        if (!string.IsNullOrWhiteSpace(annotatedOutputPath))
            writer.WriteLine($"标注副本：{annotatedOutputPath}");
        if (result.RuleSet.Warnings.Count > 0)
        {
            writer.WriteLine("规则加载警告：");
            foreach (var warning in result.RuleSet.Warnings)
                writer.WriteLine("  - " + warning);
        }

        writer.WriteLine();
        writer.WriteLine("规则覆盖表（general.yaml / 2026-body-v1.yaml 的每一条可追踪项）：");
        foreach (var row in result.Coverage)
        {
            var status = row.Implemented ? "已实现" : "未检查";
            writer.WriteLine($"  [{status}] {row.RuleId}  D2026 {row.Clause}  {row.YamlPath}  适用 {row.Applicable.ToString(CultureInfo.InvariantCulture)}，问题 {row.Failures.ToString(CultureInfo.InvariantCulture)}");
            if (!string.IsNullOrWhiteSpace(row.Reason))
                writer.WriteLine($"          原因：{row.Reason}");
        }

        writer.WriteLine();
        writer.WriteLine("不符合项：");
        if (result.Findings.Count == 0)
        {
            writer.WriteLine("  （无）");
        }
        else
        {
            for (var index = 0; index < result.Findings.Count; index++)
            {
                var finding = result.Findings[index];
                writer.WriteLine($"  [{index + 1}] {finding.RuleId}  D2026 {finding.Clause}");
                var occurrenceCount = finding.Occurrences.Count;
                writer.WriteLine(occurrenceCount > 1
                    ? $"      位置：{finding.Location}（同类共 {occurrenceCount.ToString(CultureInfo.InvariantCulture)} 处）"
                    : $"      位置：{finding.Location}");
                if (occurrenceCount > 1)
                {
                    foreach (var occurrence in finding.Occurrences.Skip(1).Take(5))
                        writer.WriteLine($"        另：{occurrence.Location}");
                    if (occurrenceCount > 6)
                        writer.WriteLine($"        另：其余 {occurrenceCount - 6} 处见 --json 的 occurrences 字段");
                }
                writer.WriteLine($"      规则位置：{finding.YamlPath}");
                writer.WriteLine($"      期望：{finding.Expected}");
                writer.WriteLine($"      实际：{finding.Actual}");
                writer.WriteLine($"      置信度：{finding.Confidence.ToString("0.00", CultureInfo.InvariantCulture)}（草案证据强度，不是合规概率）");
                writer.WriteLine($"      说明：{finding.Message}");
            }
        }

        writer.WriteLine();
        writer.WriteLine("未检查清单：");
        if (result.UncheckedItems.Count == 0)
        {
            writer.WriteLine("  （无）");
        }
        else
        {
            foreach (var item in result.UncheckedItems)
            {
                var where = string.IsNullOrWhiteSpace(item.Location) ? "规则级" : item.Location;
                writer.WriteLine($"  - {item.RuleId} @ {where}：{item.Reason}");
            }
        }

        writer.WriteLine();
        writer.WriteLine("提示：draft 规则只表示“本次可比较项未发现差异”，不代表整篇论文合规；未检查项必须保留给人工复核。");
    }

    public static string ToJson(RuleAnalysisResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var payload = new
        {
            schemaVersion = "thesis-analysis-1",
            input = result.InputPath,
            ruleSet = new
            {
                profile = result.RuleSet.ProfileId,
                version = result.RuleSet.RuleVersion,
                status = result.RuleSet.Status,
                draft = result.RuleSet.IsDraft,
                source = result.RuleSet.SourcePath,
                note = "正文 R11 字体/字号来自 S2026；其余组仍为 S1（2025版）草案值，待 2026 全文逐条复核。"
            },
            summary = new
            {
                paragraphs = result.ParagraphCount,
                checkedParagraphs = result.CheckedParagraphCount,
                passedChecks = result.PassedChecks,
                failedChecks = result.FailedChecks,
                uncheckedChecks = result.UncheckedChecks,
                findings = result.Findings.Count,
                suppressedFindings = result.SuppressedFindings,
                uncheckedItems = result.UncheckedItems.Count,
                implementedRules = result.Coverage.Count(row => row.Implemented),
                uncheckedRules = result.Coverage.Count(row => !row.Implemented)
            },
            coverage = result.Coverage.Select(row => new
            {
                ruleId = row.RuleId,
                yamlPath = row.YamlPath,
                clause = row.Clause,
                d2026 = "D2026 " + row.Clause,
                objectKind = row.ObjectKind,
                property = row.Property,
                expected = row.Expected,
                status = row.Implemented ? "已实现" : "未检查",
                implemented = row.Implemented,
                reason = row.Reason,
                applicable = row.Applicable,
                failures = row.Failures
            }),
            findings = result.Findings.Select(finding => new
            {
                ruleId = finding.RuleId,
                yamlPath = finding.YamlPath,
                clause = finding.Clause,
                d2026 = "D2026 " + finding.Clause,
                location = finding.Location,
                paragraphIndex = finding.ParagraphIndex,
                runIndex = finding.RunIndex,
                excerpt = finding.Excerpt,
                expected = finding.Expected,
                actual = finding.Actual,
                confidence = finding.Confidence,
                severity = finding.Severity,
                fixLevel = finding.FixLevel,
                occurrenceCount = finding.Occurrences.Count,
                occurrences = finding.Occurrences.Select(occurrence => new
                {
                    location = occurrence.Location,
                    paragraphIndex = occurrence.ParagraphIndex,
                    runIndex = occurrence.RunIndex,
                    excerpt = occurrence.Excerpt,
                    actual = occurrence.Actual
                }),
                message = finding.Message
            }),
            @unchecked = result.UncheckedItems.Select(item => new
            {
                ruleId = item.RuleId,
                yamlPath = item.YamlPath,
                clause = item.Clause,
                d2026 = item.Clause is null ? null : "D2026 " + item.Clause,
                location = item.Location,
                paragraphIndex = item.ParagraphIndex,
                reason = item.Reason
            })
        };
        return JsonSerializer.Serialize(payload, JsonOptions);
    }

    public static void WriteJson(RuleAnalysisResult result, TextWriter writer)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(writer);
        writer.WriteLine(ToJson(result));
    }
}
