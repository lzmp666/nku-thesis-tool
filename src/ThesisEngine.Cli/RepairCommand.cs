using ThesisEngine.Core;
using System.Globalization;
using System.Xml;

namespace ThesisEngine.Cli;

/// <summary>
/// Demo repair workflow. A user can explicitly select one executable patch or
/// opt into the complete auto batch; suggestions/manual findings stay visible.
/// </summary>
public sealed class RepairCommand(IDemoAnalysisService analysisService, TextWriter? output = null, TextWriter? error = null)
{
    public int Execute(string[] arguments, TextWriter? report = null, TextWriter? diagnostics = null)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        report ??= output ?? Console.Out;
        diagnostics ??= error ?? Console.Error;

        if (!TryParse(arguments, out var inputPath, out var outputPath, out var selection))
        {
            diagnostics.WriteLine("用法：ThesisEngine.Cli repair --demo <输入.docx> --output <修复副本.docx> (--all | --patch <编号>)");
            return 2;
        }
        if (!File.Exists(inputPath))
        {
            diagnostics.WriteLine($"输入错误：找不到文件：{inputPath}");
            return 1;
        }
        if (!string.Equals(Path.GetExtension(inputPath), ".docx", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Path.GetExtension(outputPath), ".docx", StringComparison.OrdinalIgnoreCase))
        {
            diagnostics.WriteLine("输入错误：输入和输出均必须使用 .docx 文件。");
            return 1;
        }

        try
        {
            var before = analysisService.Analyze(inputPath, DemoRuleSet.NkuDraftSample);
            var all = PatchEngine.PlanPatches(before);
            var patches = selection.All
                ? all.Patches
                : selection.Index is { } index && index >= 0 && index < all.Patches.Count
                    ? new[] { all.Patches[index] }
                    : throw new ArgumentOutOfRangeException(nameof(arguments), "补丁编号超出范围。");
            if (patches.Count == 0)
                throw new InvalidOperationException("没有可执行的 auto 补丁；suggest/manual 项需要用户单独处理。");

            var selectedPlan = new PatchPlan(patches, all.DeferredDiagnostics);
            report.WriteLine("=== 论文格式修复计划（演示）===");
            report.WriteLine($"计划补丁：{patches.Count.ToString(CultureInfo.InvariantCulture)}");
            for (var planIndex = 0; planIndex < patches.Count; planIndex++)
            {
                var patch = patches[planIndex];
                report.WriteLine($"  [{planIndex + 1}] {patch.RuleId} @ {patch.Location}");
                report.WriteLine($"      原值：{patch.ExpectedBefore}；目标值：{patch.DesiredAfter}");
                report.WriteLine($"      依据：{patch.Reason}");
            }
            var deferredCount = all.DeferredDiagnostics.Count + all.Patches.Count - patches.Count;
            report.WriteLine($"延后处理诊断：{deferredCount.ToString(CultureInfo.InvariantCulture)}");

            var applied = PatchEngine.ApplyPatches(inputPath, selectedPlan, outputPath);
            if (!applied.Succeeded)
                throw new InvalidDataException(applied.Failures[0].Reason);
            var after = analysisService.Analyze(applied.OutputPath!, DemoRuleSet.NkuDraftSample);
            var afterPlan = PatchEngine.PlanPatches(after);

            report.WriteLine("=== 论文格式修复（演示）===");
            report.WriteLine($"输入文件：{inputPath}");
            report.WriteLine($"输出文件：{applied.OutputPath}");
            report.WriteLine($"计划补丁：{patches.Count.ToString(CultureInfo.InvariantCulture)}");
            report.WriteLine($"已修复：{applied.AppliedPatches.Count.ToString(CultureInfo.InvariantCulture)}");
            report.WriteLine($"修复后诊断数：{after.Diagnostics.Count.ToString(CultureInfo.InvariantCulture)}");
            report.WriteLine($"修复后仍可自动处理：{afterPlan.Patches.Count.ToString(CultureInfo.InvariantCulture)}");
            report.WriteLine($"未处理诊断：{afterPlan.DeferredDiagnostics.Count.ToString(CultureInfo.InvariantCulture)}");
            report.WriteLine("提示：草稿 / 演示结果仅供演示和测试，不代表合规结论。");
            return 0;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException or InvalidDataException or NotSupportedException or InvalidOperationException or XmlException)
        {
            diagnostics.WriteLine($"修复失败：{exception.Message}");
            return 1;
        }
    }

    private static bool TryParse(string[] arguments, out string inputPath, out string outputPath, out Selection selection)
    {
        inputPath = string.Empty;
        outputPath = string.Empty;
        selection = default;
        if (arguments.Length < 6 ||
            !string.Equals(arguments[0], "repair", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(arguments[1], "--demo", StringComparison.Ordinal) ||
            string.IsNullOrWhiteSpace(arguments[2]) ||
            !string.Equals(arguments[3], "--output", StringComparison.Ordinal) ||
            string.IsNullOrWhiteSpace(arguments[4]))
            return false;

        inputPath = arguments[2];
        outputPath = arguments[4];
        if (arguments.Length == 6 && string.Equals(arguments[5], "--all", StringComparison.Ordinal))
        {
            selection = new Selection(true, null);
            return true;
        }
        if (arguments.Length == 7 && string.Equals(arguments[5], "--patch", StringComparison.Ordinal) &&
            int.TryParse(arguments[6], NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index > 0)
        {
            selection = new Selection(false, index - 1);
            return true;
        }
        return false;
    }

    private readonly record struct Selection(bool All, int? Index);
}
