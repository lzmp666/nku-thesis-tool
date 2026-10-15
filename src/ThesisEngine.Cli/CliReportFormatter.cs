using System.Globalization;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

public static class CliReportFormatter
{
    private static readonly IReadOnlyDictionary<string, string> UnsupportedScopeLabels =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["headers/footers"] = "页眉/页脚",
            ["footnotes"] = "脚注",
            ["endnotes"] = "尾注",
            ["tables"] = "表格",
            ["textboxes"] = "文本框",
            ["content controls"] = "内容控件",
            ["custom XML content"] = "自定义 XML 内容",
            ["altChunk content"] = "外部文档片段（altChunk）",
            ["hyperlinks"] = "超链接",
            ["smart tags"] = "智能标记",
            ["simple fields"] = "简单域",
            ["tracked changes"] = "修订内容",
            ["nested body paragraphs"] = "嵌套正文段落",
            ["word/theme/theme1.xml"] = "主题文件缺失或无法读取",
            ["theme fonts"] = "主题字体尚未解析"
        };

    public static string Format(AnalysisResult result, string? inputPath) =>
        Format(result, inputPath, null);

    public static string Format(AnalysisResult result, string? inputPath = null, string? annotatedOutputPath = null)
    {
        ArgumentNullException.ThrowIfNull(result);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        Write(result, writer, inputPath, annotatedOutputPath);
        return writer.ToString();
    }

    public static void Write(AnalysisResult result, TextWriter writer, string? inputPath = null, string? annotatedOutputPath = null)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(writer);

        var diagnostics = result.Diagnostics ?? Array.Empty<Diagnostic>();
        var unsupported = result.UnsupportedScopes ?? Array.Empty<string>();
        var count = diagnostics.Count.ToString(CultureInfo.InvariantCulture);

        writer.WriteLine("=== 论文格式检查报告 ===");
        writer.WriteLine("=== 论文格式分析（演示）===");
        writer.WriteLine("模式：草稿/演示（规则覆盖范围有限）");
        writer.WriteLine("Mode: DRAFT/DEMO (draft/demo; partial rule coverage)");
        if (!string.IsNullOrWhiteSpace(inputPath))
            writer.WriteLine($"输入文件：{inputPath}");
        writer.WriteLine("状态：分析完成");
        writer.WriteLine($"诊断数量：{count}（诊断数：{count}，发现问题：{count}）");
        writer.WriteLine($"规则版本：{ValueOrUnknown(result.RuleVersion)}");
        var coverage = ValueOrUnknown(result.Coverage);
        writer.WriteLine($"已检查范围：{coverage}（覆盖范围：{coverage}）");

        writer.WriteLine("未覆盖范围：");
        if (unsupported.Count == 0)
        {
            writer.WriteLine("  （分析服务未报告；本演示仍只覆盖部分规则）");
        }
        else
        {
            foreach (var scope in unsupported)
                writer.WriteLine($"  - {DisplayUnsupportedScope(scope)}");
        }

        var errorFindings = diagnostics
            .Where(diagnostic => string.Equals(diagnostic.Severity, "error", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (errorFindings.Length > 0)
        {
            writer.WriteLine("错误级别的问题：");
            foreach (var finding in errorFindings)
                writer.WriteLine($"  - {ValueOrUnknown(finding.RuleId)}，位置：{ValueOrUnknown(finding.Location)}；说明：{ValueOrUnknown(finding.Message)}");
        }

        writer.WriteLine("问题详情：（诊断详情）");
        if (diagnostics.Count == 0)
        {
            writer.WriteLine("  （无）");
        }
        else
        {
            for (var index = 0; index < diagnostics.Count; index++)
            {
                var diagnostic = diagnostics[index];
                writer.WriteLine($"  [{index + 1}] {DisplaySeverity(diagnostic.Severity)} {ValueOrUnknown(diagnostic.RuleId)}");
                writer.WriteLine($"      位置：{ValueOrUnknown(diagnostic.Location)}");
                writer.WriteLine($"      期望：{ValueOrUnknown(diagnostic.Expected)}");
                writer.WriteLine($"      实际：{ValueOrUnknown(diagnostic.Actual)}");
                writer.WriteLine($"      修复级别：{DisplayFixLevel(diagnostic.FixLevel)}");
                writer.WriteLine($"      说明：{ValueOrUnknown(diagnostic.Message)}");
            }
        }

        if (!string.IsNullOrWhiteSpace(annotatedOutputPath))
            writer.WriteLine($"标注文件：{annotatedOutputPath}");
        writer.WriteLine("提示：草稿 / 演示结果仅供演示和测试，不代表合规结论（不代表规范合规结论）。");
        writer.WriteLine("Notice: DRAFT/DEMO only; 仅供演示/测试，不代表合规结论。");
    }

    private static string DisplaySeverity(string? severity) =>
        severity?.ToLowerInvariant() switch
        {
            "error" => "错误",
            "warning" => "警告",
            "info" => "提示",
            _ => ValueOrUnknown(severity)
        };

    private static string DisplayFixLevel(string? fixLevel) =>
        fixLevel?.ToLowerInvariant() switch
        {
            "auto" => "自动",
            "suggest" => "建议",
            "manual" => "手动",
            "none" => "无需修复",
            _ => ValueOrUnknown(fixLevel)
        };

    private static string DisplayUnsupportedScope(string? scope) =>
        scope is not null && UnsupportedScopeLabels.TryGetValue(scope, out var label)
            ? label
            : ValueOrUnknown(scope);

    private static string ValueOrUnknown(string? value) =>
        string.IsNullOrWhiteSpace(value) || string.Equals(value, "(unknown)", StringComparison.OrdinalIgnoreCase)
            ? "（未知）"
            : value;
}
