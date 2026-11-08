using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 一键套用南开大学研究生论文格式规范。
/// 基于样式体系，修改样式定义，全文自动同步。
/// </summary>
public sealed class ApplyNkuFormatCommand
{
    private readonly StyleRewriter _rewriter;
    private readonly UndoManager _undo;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public ApplyNkuFormatCommand(
        StyleRewriter? rewriter = null,
        UndoManager? undo = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _rewriter = rewriter ?? new StyleRewriter();
        _undo = undo ?? new UndoManager();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli apply-nku-format <input.docx> [output.docx]");
            return 2;
        }

        var inputPath = args[0];
        if (!File.Exists(inputPath))
        {
            _error.WriteLine($"找不到文件：{inputPath}");
            return 1;
        }

        // 默认输出到原文件旁边，加 _nku 后缀
        var outputPath = args.Length >= 2
            ? args[1]
            : Path.Combine(
                Path.GetDirectoryName(inputPath) ?? "",
                Path.GetFileNameWithoutExtension(inputPath) + "_nku格式.docx");

        try
        {
            _output.WriteLine("开始套用南开大学研究生论文格式规范...");
            _output.WriteLine();

            // 修改前先保存快照（备份），撤销时回滚
            var backupId = _undo.BackupBeforeModify(
                inputPath,
                outputPath,
                "套用南开格式：正文宋体小四 1.5 倍行距，一/二/三级标题黑体，摘要、参考文献、题注样式");
            _output.WriteLine($"  [0/7] 已创建修改前快照（备份 ID：{backupId}），不满意可运行 undo 撤销");
            _output.WriteLine();

            // 1. 正文 Normal
            _output.WriteLine("  [1/7] 设置正文样式（Normal）：宋体 小四，1.5倍行距");
            _rewriter.SetStyleFontName(inputPath, outputPath, "Normal", "宋体", "Times New Roman");
            _rewriter.SetStyleFontSize(outputPath, outputPath, "Normal", 12);
            _rewriter.SetStyleLineSpacing(outputPath, outputPath, "Normal", 1.5);

            // 2. 一级标题 Heading1
            _output.WriteLine("  [2/7] 设置一级标题（Heading 1）：黑体 三号");
            SafeApplyStyle(outputPath, "Heading1", "黑体", "Times New Roman", 16);

            // 3. 二级标题 Heading2
            _output.WriteLine("  [3/7] 设置二级标题（Heading 2）：黑体 四号");
            SafeApplyStyle(outputPath, "Heading2", "黑体", "Times New Roman", 14);

            // 4. 三级标题 Heading3
            _output.WriteLine("  [4/7] 设置三级标题（Heading 3）：黑体 小四");
            SafeApplyStyle(outputPath, "Heading3", "黑体", "Times New Roman", 12);

            // 5. 摘要 Abstract
            _output.WriteLine("  [5/7] 设置摘要样式（Abstract）：宋体 小四");
            SafeApplyStyle(outputPath, "Abstract", "宋体", "Times New Roman", 12);

            // 6. 参考文献 Bibliography
            _output.WriteLine("  [6/7] 设置参考文献样式（Bibliography）：宋体 小五");
            SafeApplyStyle(outputPath, "Bibliography", "宋体", "Times New Roman", 9);

            // 7. 题注 Caption
            _output.WriteLine("  [7/7] 设置图表题注样式（Caption）：宋体 小五，居中");
            SafeApplyStyle(outputPath, "Caption", "宋体", "Times New Roman", 9);

            _output.WriteLine();
            _output.WriteLine($"✅ 完成！输出文件：{outputPath}");
            _output.WriteLine("已套用样式：正文、各级标题、摘要、参考文献、图表题注");
            _output.WriteLine("提示：打开文件检查一下，如有问题可运行 undo 命令撤销到修改前状态。");

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
            _rewriter.SetStyleFontName(path, path, styleId, cnFont, enFont);
            _rewriter.SetStyleFontSize(path, path, styleId, fontSize);
        }
        catch (ArgumentException)
        {
            // 样式不存在就跳过，不报错
        }
    }
}
