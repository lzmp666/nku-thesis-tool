using System.Text.Encodings.Web;
using System.Text.Json;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 摘要格式检查命令（analyze-abstract）。独立入口，只查不改。
/// 依据 general.yaml R4/R5（4.5 表4.4），规则状态 draft，提示需人工确认。
/// </summary>
public sealed class AbstractFormatAnalyzeCommand
{
    private readonly AbstractFormatAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public AbstractFormatAnalyzeCommand(
        AbstractFormatAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new AbstractFormatAnalyzer();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        string? path = null;
        var json = false;

        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--input" when i + 1 < args.Length:
                    path = args[++i];
                    break;
                case "--json":
                    json = true;
                    break;
                case "--help" or "-h":
                    WriteUsage();
                    return 0;
                default:
                    if (args[i].StartsWith('-'))
                    {
                        _error.WriteLine($"未知参数：{args[i]}");
                        WriteUsage();
                        return 2;
                    }
                    path ??= args[i];
                    break;
            }
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            WriteUsage();
            return 2;
        }
        if (!File.Exists(path))
        {
            _error.WriteLine($"找不到文件：{path}");
            return 1;
        }

        try
        {
            var result = _analyzer.Analyze(path);
            if (json)
            {
                _output.WriteLine(JsonSerializer.Serialize(result, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                }));
            }
            else
            {
                WriteReport(result);
            }
            return 0;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException or ArgumentException or System.Xml.XmlException)
        {
            _error.WriteLine($"分析失败：{ex.Message}");
            return 1;
        }
    }

    private void WriteReport(AbstractFormatResult result)
    {
        _output.WriteLine("=== 中英文摘要格式检查报告（R4/R5） ===");
        _output.WriteLine();
        _output.WriteLine($"依据：{result.SourceNote}");
        _output.WriteLine();

        if (result.Sections.Count == 0)
        {
            _output.WriteLine("未识别到任何摘要段落，报告不据此判定合规。");
        }
        else
        {
            _output.WriteLine("识别到的摘要段落：");
            foreach (var s in result.Sections)
            {
                var title = s.TitleParagraph.HasValue ? $"标题第 {s.TitleParagraph} 段" : "标题未识别";
                var body = s.BodyParagraphs.Count > 0 ? string.Join("、", s.BodyParagraphs.Select(p => $"第 {p} 段")) : "无正文段落";
                var kw = s.KeywordParagraph.HasValue ? $"第 {s.KeywordParagraph} 段" : "未识别";
                _output.WriteLine($"  - {s.Section}：{title}；正文 {body}；关键词行 {kw}");
            }
        }
        _output.WriteLine($"实际检查段落：{result.CheckedParagraphs} 段");
        _output.WriteLine();

        _output.WriteLine($"问题：{result.Issues.Count} 条");
        foreach (var issue in result.Issues)
            _output.WriteLine($"  - [{issue.Code}] {issue.Title}（第 {issue.ParagraphIndex} 段）：期望「{issue.Expected}」，实际「{issue.Actual}」");
        _output.WriteLine();

        _output.WriteLine($"未检查 / 未判定：{result.NotChecked.Count} 项");
        foreach (var item in result.NotChecked)
            _output.WriteLine($"  - {item}");

        if (result.MissingParts.Count > 0)
        {
            _output.WriteLine();
            _output.WriteLine($"缺少部件：{string.Join("、", result.MissingParts)}");
            _output.WriteLine("（缺部件不等于合规，相关判据已列入未检查项）");
        }
    }

    private void WriteUsage()
    {
        _output.WriteLine("用法：ThesisEngine.Cli analyze-abstract <file.docx> [--json]");
        _output.WriteLine("检查中英文摘要标题/正文的字体、字号、加粗、对齐、行距与段距（general.yaml R4/R5，draft）。");
        _output.WriteLine("只查不改；样式链判不出的维度列入「未判定」，不按合规处理。");
    }
}
