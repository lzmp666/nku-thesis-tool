using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 图表排版分析命令。
/// </summary>
public sealed class FigureTableAnalyzeCommand
{
    private readonly FigureTableAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public FigureTableAnalyzeCommand(
        FigureTableAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new FigureTableAnalyzer();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli analyze-figures <file.docx>");
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

            _output.WriteLine("=== 图表排版分析报告 ===");
            _output.WriteLine();
            _output.WriteLine($"图片数量：{result.ImagesCount} 张");
            _output.WriteLine($"表格数量：{result.TablesCount} 个");
            _output.WriteLine();

            if (result.TableIssues.Count > 0)
            {
                _output.WriteLine($"⚠️ 表格问题：{result.TableIssues.Count} 个");
                foreach (var issue in result.TableIssues)
                {
                    _output.WriteLine($"  - {issue}");
                }
                _output.WriteLine();
            }

            if (result.PageBreakIssues.Count > 0)
            {
                _output.WriteLine($"⚠️ 跨页问题：{result.PageBreakIssues.Count} 处");
                foreach (var issue in result.PageBreakIssues.Take(10))
                {
                    _output.WriteLine($"  - {issue}");
                }
                if (result.PageBreakIssues.Count > 10)
                {
                    _output.WriteLine($"  ... 还有 {result.PageBreakIssues.Count - 10} 处");
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"分析失败：{ex.Message}");
            return 1;
        }
    }
}
