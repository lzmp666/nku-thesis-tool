using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 一键应用图表排版格式。
/// </summary>
public sealed class ApplyFigureFormatCommand
{
    private readonly FigureTableRewriter _rewriter;
    private readonly UndoManager _undo;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public ApplyFigureFormatCommand(
        FigureTableRewriter? rewriter = null,
        UndoManager? undo = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _rewriter = rewriter ?? new FigureTableRewriter();
        _undo = undo ?? new UndoManager();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli apply-figure-format <input.docx> [output.docx]");
            return 2;
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
                Path.GetFileNameWithoutExtension(inputPath) + "_图表格式.docx");

        try
        {
            _output.WriteLine("开始应用图表排版格式...");
            _output.WriteLine();

            // 修改前先保存快照（备份），撤销时回滚
            var backupId = _undo.BackupBeforeModify(
                inputPath,
                outputPath,
                "图表排版：所有表格改为三线表，所有图片居中对齐");
            _output.WriteLine($"  [0/2] 已创建修改前快照（备份 ID：{backupId}），不满意可运行 undo 撤销");
            _output.WriteLine();

            _output.WriteLine("  [1/2] 将所有表格改为三线表");
            _rewriter.ApplyThreeLineTables(inputPath, outputPath);

            _output.WriteLine("  [2/2] 将所有图片居中对齐");
            _rewriter.CenterAlignImages(outputPath, outputPath);

            _output.WriteLine();
            _output.WriteLine($"✅ 完成！输出文件：{outputPath}");
            _output.WriteLine("提示：如有问题可运行 undo 命令撤销到修改前状态。");

            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"失败：{ex.Message}");
            return 1;
        }
    }
}
