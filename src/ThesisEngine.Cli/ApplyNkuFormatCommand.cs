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
            _output.WriteLine("开始套用南开研究生论文格式...");
            _output.WriteLine();

            // 修改 Normal 样式（正文）：宋体小四，西文 Times New Roman
            _output.WriteLine("  [1/4] 设置正文样式（Normal）：宋体 小四");
            _rewriter.SetStyleFontName(inputPath, outputPath, "Normal", "宋体", "Times New Roman");
            _rewriter.SetStyleFontSize(outputPath, outputPath, "Normal", 12);

            // 修改 Heading1：黑体三号
            _output.WriteLine("  [2/4] 设置一级标题（Heading 1）：黑体 三号");
            _rewriter.SetStyleFontName(outputPath, outputPath, "Heading1", "黑体", "Times New Roman");
            _rewriter.SetStyleFontSize(outputPath, outputPath, "Heading1", 16);

            // 修改 Heading2：黑体四号
            _output.WriteLine("  [3/4] 设置二级标题（Heading 2）：黑体 四号");
            _rewriter.SetStyleFontName(outputPath, outputPath, "Heading2", "黑体", "Times New Roman");
            _rewriter.SetStyleFontSize(outputPath, outputPath, "Heading2", 14);

            // 修改 Heading3：黑体小四
            _output.WriteLine("  [4/4] 设置三级标题（Heading 3）：黑体 小四");
            _rewriter.SetStyleFontName(outputPath, outputPath, "Heading3", "黑体", "Times New Roman");
            _rewriter.SetStyleFontSize(outputPath, outputPath, "Heading3", 12);

            _output.WriteLine();
            _output.WriteLine($"✅ 完成！输出文件：{outputPath}");
            _output.WriteLine("提示：打开文件检查一下，确认格式是否符合要求。");

            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"失败：{ex.Message}");
            return 1;
        }
    }
}
