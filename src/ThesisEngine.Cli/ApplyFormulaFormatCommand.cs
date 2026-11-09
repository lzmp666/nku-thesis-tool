using ThesisEngine.Core;
using ThesisEngine.Core.Formula;

namespace ThesisEngine.Cli;

/// <summary>
/// 公式格式写入命令（#114 第二档：只改格式，不碰内容）。
/// 写完立刻复诊一遍：改前几条、改后几条，剩几条为什么没改，都得看见。
/// </summary>
public sealed class ApplyFormulaFormatCommand
{
    private readonly FormulaFormatter _formatter;
    private readonly FormulaAnalyzer _analyzer;
    private readonly UndoManager _undo;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public ApplyFormulaFormatCommand(
        FormulaFormatter? formatter = null,
        FormulaAnalyzer? analyzer = null,
        UndoManager? undo = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _formatter = formatter ?? new FormulaFormatter();
        _analyzer = analyzer ?? new FormulaAnalyzer();
        _undo = undo ?? new UndoManager();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli apply-formula-format <input.docx> [output.docx]");
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
                Path.GetFileNameWithoutExtension(inputPath) + "_公式格式.docx");

        try
        {
            _output.WriteLine("=== 公式格式写入 ===");
            _output.WriteLine();
            _output.WriteLine($"输入：{inputPath}");

            var before = _analyzer.Analyze(inputPath);

            var backupId = _undo.BackupBeforeModify(
                inputPath,
                outputPath,
                "公式格式：居中、式号右对齐、字体 Cambria Math、段前段后归零");
            _output.WriteLine($"备份 ID：{backupId}（不满意可运行 undo 撤销）");
            _output.WriteLine();

            var result = _formatter.Apply(inputPath, outputPath);

            _output.WriteLine(
                $"独立公式 {result.StandaloneCount} 个｜行内公式 {result.InlineCount} 个（整段跳过）｜" +
                $"表格排版 {result.TableRowCount} 个（跳过）");
            _output.WriteLine();

            if (result.Fixes.Count > 0)
            {
                _output.WriteLine($"--- 已修改 {result.Fixes.Count} 处 ---");
                foreach (var fix in result.Fixes)
                {
                    _output.WriteLine(
                        $"  段落 {fix.ParagraphIndex} / 公式 {fix.FormulaIndex}  [{fix.Code}] {fix.Description}");
                }
                _output.WriteLine();
            }
            else
            {
                _output.WriteLine("没有需要修改的格式项。");
                _output.WriteLine();
            }

            if (result.Skips.Count > 0)
            {
                _output.WriteLine($"--- 没有自动改的 {result.Skips.Count} 处 ---");
                foreach (var skip in result.Skips)
                {
                    _output.WriteLine($"  - {skip}");
                }
                _output.WriteLine();
            }

            WriteFollowUp(before, outputPath);
            WriteUncheckedItems();

            _output.WriteLine($"输出：{outputPath}");
            _output.WriteLine("提示：打开检查一下；不满意运行 undo 撤销到修改前状态。");
            return 0;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException or ArgumentException)
        {
            _error.WriteLine($"写入失败：{ex.Message}");
            return 1;
        }
    }

    /// <summary>复诊：改前改后各报一次数，剩下的逐条列出来，别让人觉得修完了。</summary>
    private void WriteFollowUp(FormulaAnalysisResult before, string outputPath)
    {
        var after = _analyzer.Analyze(outputPath);

        var beforeClause = before.ClauseIssues.Count();
        var beforeConvention = before.ConventionIssues.Count();
        var afterClause = after.ClauseIssues.Count();
        var afterConvention = after.ConventionIssues.Count();

        _output.WriteLine("--- 复诊 ---");
        _output.WriteLine($"修改前：不符合规范 {beforeClause} 处、排版惯例 {beforeConvention} 处");
        _output.WriteLine($"修改后：不符合规范 {afterClause} 处、排版惯例 {afterConvention} 处");

        var remaining = after.Issues.ToList();
        if (remaining.Count == 0)
        {
            _output.WriteLine("  复诊未发现问题。");
            _output.WriteLine();
            return;
        }

        _output.WriteLine("  剩余（需人工处理）：");
        foreach (var issue in remaining)
        {
            _output.WriteLine(
                $"    [{issue.Code}] {issue.Title}｜段落 {issue.ParagraphIndex}" +
                (issue.FormulaIndex.HasValue ? $" / 公式 {issue.FormulaIndex}" : string.Empty) +
                $"｜当前：{issue.Current}");
        }
        _output.WriteLine();
        _output.WriteLine("  注意：复诊为 0 只说明 F1–F5 这几项没发现差异，不等于公式排版整体合规。");
        _output.WriteLine();
    }

    private void WriteUncheckedItems()
    {
        _output.WriteLine("--- 本次未处理 ---");
        _output.WriteLine("  - 式号缺圆括号：加括号要改文字，本档不做");
        _output.WriteLine("  - 式号字号写在公式对象里（m:rPr）：动它会改公式本体排版，留给人工");
        _output.WriteLine("  - 行内公式：整段不改，它只是句子的一部分");
        _output.WriteLine("  - 两列表格排版：位置由列宽决定，改段落对齐看不出效果，留给人工");
        _output.WriteLine("  - 公式编号连续性与章节重排：#114 明确不做");
        _output.WriteLine("  - 未显式设置字体/字号的 run 视为继承样式，不覆盖（归样式检查）");
        _output.WriteLine();
    }
}
