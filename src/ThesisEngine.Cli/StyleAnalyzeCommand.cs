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
            foreach (var (id, name) in result.AllStyles)
            {
                _output.WriteLine($"  - {name} ({id})");
            }
            _output.WriteLine();

            var tableParagraphCount = result.StyledParagraphs.Count(p => p.InTable);
            _output.WriteLine($"套用样式的段落：{result.StyledParagraphs.Count} 段" +
                (tableParagraphCount > 0 ? $"（其中表格内 {tableParagraphCount} 段）" : string.Empty));
            _output.WriteLine();

            if (result.InvalidStyleParagraphs.Count > 0)
            {
                _output.WriteLine($"⚠️ 引用了样式表中不存在的样式：{result.InvalidStyleParagraphs.Count} 段");
                _output.WriteLine("（这些段落的样式引用是无效的，Word 会按默认格式渲染）");
                foreach (var para in result.InvalidStyleParagraphs.Take(10))
                {
                    var where = para.InTable ? "（表格内）" : string.Empty;
                    _output.WriteLine($"  - \"{para.Preview}\"  引用样式：{para.StyleId} {where}");
                }
                if (result.InvalidStyleParagraphs.Count > 10)
                {
                    _output.WriteLine($"  ... 还有 {result.InvalidStyleParagraphs.Count - 10} 段");
                }
                _output.WriteLine();
            }

            _output.WriteLine($"⚠️ 没有套用样式的段落：{result.UnstyledParagraphs.Count} 段");
            _output.WriteLine("（这些是手动打的格式，建议套上对应样式）");
            foreach (var para in result.UnstyledParagraphs.Take(10))
            {
                var where = para.InTable ? "（表格内）" : string.Empty;
                _output.WriteLine($"  - \"{para.Preview}\"  [{para.DirectFormat}]{where}");
            }
            if (result.UnstyledParagraphs.Count > 10)
            {
                _output.WriteLine($"  ... 还有 {result.UnstyledParagraphs.Count - 10} 段");
            }

            _output.WriteLine();

            if (result.DirectFormatOverrides.Count > 0)
            {
                _output.WriteLine($"⚠️ 套了样式、但格式仍被手动设置压住：{result.DirectFormatOverrides.Count} 段");
                _output.WriteLine("（Word 里手动格式优先级高于样式，这些段落的字体字号不会跟着样式变）");
                foreach (var para in result.DirectFormatOverrides.Take(10))
                {
                    var where = para.InTable ? "（表格内）" : string.Empty;
                    _output.WriteLine($"  - \"{para.Preview}\"  样式：{para.StyleName}  [{para.Overrides}]{where}");
                }
                if (result.DirectFormatOverrides.Count > 10)
                {
                    _output.WriteLine($"  ... 还有 {result.DirectFormatOverrides.Count - 10} 段");
                }
            }
            else
            {
                _output.WriteLine("✅ 套用样式的段落没有被直接格式覆盖，改样式会真正生效。");
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
