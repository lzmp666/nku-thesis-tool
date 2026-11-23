using System.Text;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// draft 规则的保守自动修正命令（apply-draft-format）。
///
/// 只把六个 draft 分析器「已明确判定」的问题段落对齐到规则的 Expected 值；
/// 分析器列在未判定的维度不进 Issue，因此绝不会被自动改动。
/// 规则状态 draft，输出明确声明「不构成规范合规结论」，并给出撤销提示。
/// </summary>
public sealed class ApplyDraftFormatCommand
{
    private readonly DraftFormatRewriter _rewriter;
    private readonly UndoManager _undo;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public ApplyDraftFormatCommand(
        DraftFormatRewriter? rewriter = null,
        UndoManager? undo = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _rewriter = rewriter ?? new DraftFormatRewriter();
        _undo = undo ?? new UndoManager();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0 || args[0] is "--help" or "-h")
        {
            WriteUsage();
            return args.Length == 0 ? 2 : 0;
        }

        var inputPath = args[0];
        if (!File.Exists(inputPath))
        {
            _error.WriteLine($"找不到文件：{inputPath}");
            return 1;
        }

        var outputPath = args.Length >= 2
            ? args[1]
            : Path.Combine(
                Path.GetDirectoryName(inputPath) ?? "",
                Path.GetFileNameWithoutExtension(inputPath) + "_draft修正.docx");

        try
        {
            _output.WriteLine("=== draft 规则保守修正（apply-draft-format） ===");
            _output.WriteLine();
            _output.WriteLine("声明：规则状态 draft，仅把已明确判定的维度机械对齐到规则期望值；");
            _output.WriteLine("      未判定维度不改动。本结果不构成规范合规结论，需人工复核。");
            _output.WriteLine();

            var backupId = _undo.BackupBeforeModify(
                inputPath,
                outputPath,
                "draft 规则保守修正：仅改六个 draft 分析器明确判定的维度");
            _output.WriteLine($"已创建修改前快照（备份 ID：{backupId}），可用 undo 撤销。");

            var result = _rewriter.Apply(inputPath, outputPath);

            _output.WriteLine($"已修正维度：{result.Applied} 处，涉及段落：{string.Join("、", result.ChangedParagraphs.Count > 0 ? result.ChangedParagraphs : new[] { "无" })}");
            if (result.Skipped.Count > 0)
            {
                _output.WriteLine($"跳过（无法安全映射）：{result.Skipped.Count} 项");
                foreach (var item in result.Skipped) _output.WriteLine($"  - {item}");
            }
            _output.WriteLine();
            _output.WriteLine(result.HasChanges
                ? $"完成！输出文件：{outputPath}"
                : "没有需要修正的明确问题，输出与输入内容一致。");
            _output.WriteLine("提示：请重新运行 report 复核，并在 Word 中实机确认。");

            return 0;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException
            or UnauthorizedAccessException or ArgumentException or System.Xml.XmlException)
        {
            _error.WriteLine($"修正失败：{ex.Message}");
            return 1;
        }
    }

    private void WriteUsage()
    {
        var usage = new StringBuilder();
        usage.AppendLine("用法：ThesisEngine.Cli apply-draft-format <input.docx> [output.docx]");
        usage.AppendLine("把六个 draft 分析器已明确判定的格式问题，保守地对齐到规则期望值。");
        usage.AppendLine("只改明确维度；未判定维度跳过。规则状态 draft，不构成合规结论。");
        usage.AppendLine("默认输出 <原名>_draft修正.docx，可用 undo 撤销。");
        _output.WriteLine(usage.ToString().TrimEnd());
    }
}
