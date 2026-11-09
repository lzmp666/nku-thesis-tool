using ThesisEngine.Core.Formula;

namespace ThesisEngine.Cli;

/// <summary>
/// 公式排版分析命令（#114 第一档：只查不改）。
/// </summary>
public sealed class FormulaAnalyzeCommand
{
    private readonly FormulaAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public FormulaAnalyzeCommand(
        FormulaAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new FormulaAnalyzer();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli analyze-formula <file.docx>");
            return 2;
        }

        var path = args[0];
        if (!File.Exists(path))
        {
            _error.WriteLine($"找不到文件：{path}");
            return 1;
        }

        try
        {
            var result = _analyzer.Analyze(path);

            _output.WriteLine("=== 公式排版分析报告 ===");
            _output.WriteLine();
            _output.WriteLine($"文件：{path}");
            _output.WriteLine($"段落数：{result.ParagraphCount}");
            _output.WriteLine($"独立公式：{result.Blocks.Count} 个（其中带式号 {result.Blocks.Count(b => b.HasEquationNumber)} 个）");
            _output.WriteLine($"行内公式：{result.InlineCount} 个（不参与排版检查，只计数）");
            _output.WriteLine();

            if (result.Blocks.Count > 0)
            {
                _output.WriteLine("--- 公式清单 ---");
                foreach (var block in result.Blocks)
                {
                    var layout = block.Layout switch
                    {
                        FormulaLayout.Standalone => "独立成段",
                        FormulaLayout.TableRow => "两列表格",
                        _ => "行内",
                    };
                    var number = block.EquationNumber == null ? "无式号" : $"式号 {block.EquationNumber}";
                    _output.WriteLine($"公式 {block.Index}（段落 {block.ParagraphIndex}）：{block.Text}");
                    _output.WriteLine($"  {layout}，{number}");
                }
                _output.WriteLine();
            }

            var clauseIssues = result.ClauseIssues.ToList();
            var conventionIssues = result.ConventionIssues.ToList();

            if (clauseIssues.Count > 0)
            {
                _output.WriteLine($"⚠️ 不符合规范：{clauseIssues.Count} 处");
                foreach (var issue in clauseIssues)
                {
                    WriteIssue(issue);
                }
                _output.WriteLine();
            }

            if (conventionIssues.Count > 0)
            {
                _output.WriteLine($"排版惯例提示：{conventionIssues.Count} 处（规范未规定，不计入规范符合性）");
                foreach (var issue in conventionIssues)
                {
                    WriteIssue(issue);
                }
                _output.WriteLine();
            }

            if (clauseIssues.Count == 0 && conventionIssues.Count == 0)
            {
                _output.WriteLine(result.Blocks.Count == 0
                    ? "文档中没有独立公式段，未产生检查结论。"
                    : "已检查的公式项均符合。");
                _output.WriteLine();
            }

            WriteUncheckedItems();

            return 0;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
        {
            _error.WriteLine($"分析失败：{ex.Message}");
            return 1;
        }
    }

    private void WriteIssue(FormulaIssue issue)
    {
        _output.WriteLine($"  [{issue.Code}] {issue.Title}");
        _output.WriteLine($"    位置：段落 {issue.ParagraphIndex}" +
            (issue.FormulaIndex.HasValue ? $" / 公式 {issue.FormulaIndex}" : string.Empty));
        _output.WriteLine($"    预览：{issue.Preview}");
        _output.WriteLine($"    当前：{issue.Current}");
        _output.WriteLine($"    应为：{issue.Expected}");
        _output.WriteLine($"    依据：{issue.Basis}");
    }

    /// <summary>
    /// 未检查项必须写出来。0 条问题分两种：查了没问题，和压根没查。
    /// 不写清楚，用户会把后者当成前者。
    /// </summary>
    private void WriteUncheckedItems()
    {
        _output.WriteLine("--- 本次未检查 ---");
        _output.WriteLine("  - 公式行距：R11 允许「段落中有数学表达式时按表达需要设置行距」，不作检查");
        _output.WriteLine("  - 公式编号连续性与章节重排：#114 明确不做，跨章节改号风险大于收益");
        _output.WriteLine("  - LaTeX 转 OMML：本档不做，只查已有公式的格式");
        _output.WriteLine("  - 两列表格排版下的公式居中：位置由列宽决定，不按段落对齐判断");
        _output.WriteLine("  - 未显式设置字体/字号的 run 视为继承样式，不在这里判对错（归样式检查）");
    }
}
