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

        // 中途失败时把这次新建的输出删掉，别留一个只做了一半的文件让人误以为完成了。
        var outputExistedBefore = File.Exists(outputPath);

        try
        {
            _output.WriteLine("=" + new string('=', 50));
            _output.WriteLine("  南开大学研究生论文 一键完整格式化");
            _output.WriteLine("=" + new string('=', 50));
            _output.WriteLine();

            // 0. 修改前先保存快照（备份），撤销时回滚
            var changeDescription =
                "一键完整格式化：自动套用段落样式、" +
                "修改样式定义（正文宋体小四 1.5 倍行距、各级标题黑体）、" +
                "表格改三线表、图片居中、修正跨页问题";
            var backupId = _undo.BackupBeforeModify(inputPath, outputPath, changeDescription);
            _output.WriteLine($"【0/5】已创建修改前快照（备份 ID：{backupId}）");
            _output.WriteLine($"  - 如对结果不满意，可运行 undo 命令一键撤销：{outputPath}");
            _output.WriteLine();

            // 1. 自动给无样式段落套样式
            _output.WriteLine("【1/5】自动识别并套用段落样式");
            _output.WriteLine("  - 检测无样式段落（含表格内段落）");
            _output.WriteLine("  - 识别章/节/小节标题、摘要、参考文献等结构");
            _output.WriteLine("  - 其他段落套用正文样式");
            var autoAssigner = new StyleAutoAssigner(_output);
            var assignment = autoAssigner.AutoAssignStyles(inputPath, outputPath);
            if (assignment.TotalAssigned == 0)
            {
                _output.WriteLine("所有段落都已有样式归属，无需调整。");
            }
            _output.WriteLine();

            // 2. 修改样式定义
            _output.WriteLine("【2/5】修改样式定义");
            _output.WriteLine("  - 正文 Normal：宋体 小四，1.5倍行距");
            _output.WriteLine("  - 一级标题：黑体 三号");
            _output.WriteLine("  - 二级标题：黑体 四号");
            _output.WriteLine("  - 三级标题：黑体 小四");
            _output.WriteLine("  - 摘要、参考文献、题注样式");

            // 正文样式是全局基础，缺了整条链都会失真，所以和其它样式一样走容错，
            // 只把缺的那几项记下来，不要因为一个样式名对不上就中止整条命令。
            var skippedStyles = new List<string>();

            void ApplyStyle(string styleId, string cnFont, double fontSize, double? lineSpacing = null)
            {
                if (!SafeApplyStyle(outputPath, styleId, cnFont, "Times New Roman", fontSize, lineSpacing))
                {
                    skippedStyles.Add(styleId);
                }
            }

            ApplyStyle("Normal", "宋体", 12, 1.5);
            ApplyStyle("Heading1", "黑体", 16);
            ApplyStyle("Heading2", "黑体", 14);
            ApplyStyle("Heading3", "黑体", 12);
            ApplyStyle("Abstract", "宋体", 12);
            ApplyStyle("Bibliography", "宋体", 9);
            ApplyStyle("Caption", "宋体", 9);

            if (skippedStyles.Count > 0)
            {
                _output.WriteLine($"  ⚠️ 样式表中没有 {string.Join("、", skippedStyles)}，这几项保持原样");
            }

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
            if (skippedStyles.Count > 0)
            {
                _output.WriteLine($"完成，{skippedStyles.Count} 项跳过（{string.Join("、", skippedStyles)}）。");
                _output.WriteLine("  这几项在文档样式表里不存在，不是改失败，是这篇文档里没有对应的样式。");
            }
            else
            {
                _output.WriteLine("✅ 全部格式化完成！");
            }
            _output.WriteLine("提示：打开文件检查一下，如有问题可运行 undo 命令撤销。");
            _output.WriteLine($"      （备份 ID：{backupId}，可用 undo --list 查看全部备份）");

            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"失败：{ex.Message}");

            // 这次才生成的文件属于半成品，删掉；原本就存在的文件不动。
            if (!outputExistedBefore && File.Exists(outputPath))
            {
                try
                {
                    File.Delete(outputPath);
                    _error.WriteLine($"已删除未完成的输出文件：{outputPath}");
                }
                catch (IOException deleteError)
                {
                    _error.WriteLine($"未完成的输出文件删除失败（{deleteError.Message}），请手动确认：{outputPath}");
                }
            }

            return 1;
        }
    }

    /// <summary>
    /// 套用一组样式属性；样式不存在时跳过并返回 false，由调用方汇总告知用户。
    /// </summary>
    private bool SafeApplyStyle(string path, string styleId, string cnFont, string enFont, double fontSize, double? lineSpacing = null)
    {
        try
        {
            _styleRewriter.SetStyleFontName(path, path, styleId, cnFont, enFont);
            _styleRewriter.SetStyleFontSize(path, path, styleId, fontSize);
            if (lineSpacing.HasValue)
            {
                _styleRewriter.SetStyleLineSpacing(path, path, styleId, lineSpacing.Value);
            }

            return true;
        }
        catch (ArgumentException)
        {
            // 样式不存在就跳过
            return false;
        }
    }
}
