using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 一键完整格式化：自动备份 + 套用样式 + 修正图表。
/// </summary>
public sealed class FullFormatCommand
{
    private readonly StyleRewriter _styleRewriter;
    private readonly FigureTableRewriter _figureRewriter;
    private readonly UndoManager _undo;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public FullFormatCommand(
        StyleRewriter? styleRewriter = null,
        FigureTableRewriter? figureRewriter = null,
        UndoManager? undo = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _styleRewriter = styleRewriter ?? new StyleRewriter();
        _figureRewriter = figureRewriter ?? new FigureTableRewriter();
        _undo = undo ?? new UndoManager();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli full-format <input.docx> [output.docx]");
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
                Path.GetFileNameWithoutExtension(inputPath) + "_南开格式完整版.docx");

        try
        {
            _output.WriteLine("=" + new string('=', 50));
            _output.WriteLine("  南开大学研究生论文 一键完整格式化");
            _output.WriteLine("=" + new string('=', 50));
            _output.WriteLine();

            // 0. 先做备份，支持撤销
            _output.WriteLine("【0/5】备份原文件（支持后续撤销）");
            var undoManager = new UndoManager();
            // 把输入文件备份成输出文件的名字，这样 undo 输出文件的时候能找到备份
            var backupId = undoManager.BackupBeforeModify(inputPath);
            // 重命名备份文件，让它和输出文件同名，这样 undo 才能找到
            var backupDir = Path.Combine(Path.GetTempPath(), "nku-thesis-tool-backups");
            var oldBackupPath = Path.Combine(backupDir, $"{backupId}_{Path.GetFileName(inputPath)}");
            var newBackupPath = Path.Combine(backupDir, $"{backupId}_{Path.GetFileName(outputPath)}");
            if (File.Exists(oldBackupPath))
            {
                File.Move(oldBackupPath, newBackupPath);
            }
            _output.WriteLine("  - 备份完成，后续可运行 undo 撤销");
            _output.WriteLine();

            // 1. 自动给无样式段落套样式
            _output.WriteLine("【1/5】自动识别并套用段落样式");
            _output.WriteLine("  - 检测无样式段落");
            _output.WriteLine("  - 短文本加粗段落识别为标题");
            _output.WriteLine("  - 其他段落套用正文样式");
            var autoAssigner = new StyleAutoAssigner();
            autoAssigner.AutoAssignStyles(inputPath, outputPath);
            _output.WriteLine();

            // 2. 修改样式定义
            _output.WriteLine("【2/5】修改样式定义");
            _output.WriteLine("  - 正文 Normal：宋体 小四，1.5倍行距");
            _output.WriteLine("  - 一级标题：黑体 三号");
            _output.WriteLine("  - 二级标题：黑体 四号");
            _output.WriteLine("  - 三级标题：黑体 小四");
            _output.WriteLine("  - 摘要、参考文献、题注样式");

            _styleRewriter.SetStyleFontName(outputPath, outputPath, "Normal", "宋体", "Times New Roman");
            _styleRewriter.SetStyleFontSize(outputPath, outputPath, "Normal", 12);
            _styleRewriter.SetStyleLineSpacing(outputPath, outputPath, "Normal", 1.5);

            SafeApplyStyle(outputPath, "Heading1", "黑体", "Times New Roman", 16);
            SafeApplyStyle(outputPath, "Heading2", "黑体", "Times New Roman", 14);
            SafeApplyStyle(outputPath, "Heading3", "黑体", "Times New Roman", 12);
            SafeApplyStyle(outputPath, "Abstract", "宋体", "Times New Roman", 12);
            SafeApplyStyle(outputPath, "Bibliography", "宋体", "Times New Roman", 9);
            SafeApplyStyle(outputPath, "Caption", "宋体", "Times New Roman", 9);
            _output.WriteLine();

            // 3. 修正图表格式
            _output.WriteLine("【3/5】修正图表格式");
            _output.WriteLine("  - 所有表格改为三线表");
            _output.WriteLine("  - 所有图片居中对齐");
            _figureRewriter.ApplyThreeLineTables(outputPath, outputPath);
            _figureRewriter.CenterAlignImages(outputPath, outputPath);
            _output.WriteLine();

            // 4. 修正跨页问题
            _output.WriteLine("【4/5】修正跨页问题");
            _output.WriteLine("  - 孤行控制，避免段落最后一行单独在一页");
            _output.WriteLine("  - 与下段同页，避免标题和正文分开");
            _figureRewriter.FixPageBreaks(outputPath, outputPath);
            _output.WriteLine();

            // 5. 完成
            _output.WriteLine("【5/5】完成");
            _output.WriteLine($"  - 输出：{outputPath}");
            _output.WriteLine();
            _output.WriteLine("✅ 全部格式化完成！");
            _output.WriteLine("提示：打开文件检查一下，如有问题可运行 undo 命令撤销。");

            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"失败：{ex.Message}");
            return 1;
        }
    }

    private void SafeApplyStyle(string path, string styleId, string cnFont, string enFont, double fontSize)
    {
        try
        {
            _styleRewriter.SetStyleFontName(path, path, styleId, cnFont, enFont);
            _styleRewriter.SetStyleFontSize(path, path, styleId, fontSize);
        }
        catch (ArgumentException)
        {
            // 样式不存在就跳过
        }
    }
}
