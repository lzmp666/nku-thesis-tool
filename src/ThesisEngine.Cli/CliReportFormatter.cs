using System.Globalization;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

public static class CliReportFormatter
{
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

        writer.WriteLine("=== 论文格式分析（演示）===");
        writer.WriteLine("模式：草稿 / 演示，仅覆盖部分格式规则");
        writer.WriteLine("Mode: DRAFT/DEMO (draft/demo; partial rule coverage)");
        if (!string.IsNullOrWhiteSpace(inputPath)) writer.WriteLine($"输入文件：{inputPath}");
        writer.WriteLine("状态：分析完成");
        writer.WriteLine($"诊断数：{diagnostics.Count.ToString(CultureInfo.InvariantCulture)}");
        writer.WriteLine($"Diagnostics: {diagnostics.Count.ToString(CultureInfo.InvariantCulture)}");
        writer.WriteLine($"规则版本：{ValueOrUnknown(result.RuleVersion)}");
        writer.WriteLine($"Rule version: {ValueOrUnknown(result.RuleVersion)}");
        writer.WriteLine($"覆盖范围：{ValueOrUnknown(result.Coverage)}");
        writer.WriteLine($"Coverage: {ValueOrUnknown(result.Coverage)}");

        writer.WriteLine("未覆盖范围：");
        writer.WriteLine("Unsupported scopes:");
        if (unsupported.Count == 0)
            writer.WriteLine("  （分析器未报告其他未覆盖范围；演示模式仍只检查部分规则）");
        else
            foreach (var scope in unsupported) writer.WriteLine($"  - {ValueOrUnknown(scope)}");

        writer.WriteLine("诊断详情：");
        var errorFindings = diagnostics
            .Where(diagnostic => string.Equals(diagnostic.Severity, "error", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (errorFindings.Length > 0)
        {
            writer.WriteLine("Error-severity findings:");
            foreach (var finding in errorFindings)
                writer.WriteLine($"  - {ValueOrUnknown(finding.RuleId)} at {ValueOrUnknown(finding.Location)}: {ValueOrUnknown(finding.Message)}");
        }
        if (diagnostics.Count == 0)
            writer.WriteLine("  （无）");
        else
            for (var index = 0; index < diagnostics.Count; index++)
            {
                var diagnostic = diagnostics[index];
                writer.WriteLine($"  [{index + 1}] {ValueOrUnknown(diagnostic.Severity)} {ValueOrUnknown(diagnostic.RuleId)}");
                writer.WriteLine($"      位置：{ValueOrUnknown(diagnostic.Location)}");
                writer.WriteLine($"      期望值：{ValueOrUnknown(diagnostic.Expected)}；实际值：{ValueOrUnknown(diagnostic.Actual)}");
                writer.WriteLine($"      处理级别：{ValueOrUnknown(diagnostic.FixLevel)}");
                writer.WriteLine($"      说明：{ValueOrUnknown(diagnostic.Message)}");
            }

        if (!string.IsNullOrWhiteSpace(annotatedOutputPath))
            writer.WriteLine($"标注文件：{annotatedOutputPath}");
        writer.WriteLine("提示：草稿 / 演示结果仅供演示和测试，不代表合规结论。");
        writer.WriteLine("Notice: DRAFT/DEMO only; 仅供演示/测试，不代表合规结论。");
    }

    private static string ValueOrUnknown(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "（未知）" : value;
}
