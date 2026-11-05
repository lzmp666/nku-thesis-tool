using ThesisEngine.Core;

namespace ThesisEngine.Cli;

public sealed class StyleAnalyzeCommand
{
    private readonly StyleAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public StyleAnalyzeCommand(
        StyleAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new StyleAnalyzer();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli analyze-styles <file.docx>");
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

            _output.WriteLine("=== 文档样式分析报告 ===");
            _output.WriteLine();

            _output.WriteLine($"共发现 {result.AllStyles.Count} 个样式：");
            foreach (var (id, name) in result.AllStyles.Take(10))
            {
                _output.WriteLine($"  - {name} ({id})");
            }
            if (result.AllStyles.Count > 10)
            {
                _output.WriteLine($"  ... 还有 {result.AllStyles.Count - 10} 个");
            }
            _output.WriteLine();

            _output.WriteLine($"套用样式的段落：{result.StyledParagraphs.Count} 段");
            _output.WriteLine();

            _output.WriteLine($"⚠️ 没有套用样式的段落：{result.UnstyledParagraphs.Count} 段");
            _output.WriteLine("（这些是手动打的格式，建议套上对应样式）");
            foreach (var para in result.UnstyledParagraphs.Take(10))
            {
                _output.WriteLine($"  - \"{para.Preview}\"  [{para.DirectFormat}]");
            }
            if (result.UnstyledParagraphs.Count > 10)
            {
                _output.WriteLine($"  ... 还有 {result.UnstyledParagraphs.Count - 10} 段");
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
