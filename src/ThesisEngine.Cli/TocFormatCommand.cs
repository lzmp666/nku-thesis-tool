using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 目录格式检查命令（analyze-toc）。R6，只查不改。
/// 目录标题黑体 16 磅加粗居中；章/一级/二级条目宋体 14/12/10.5 磅、缩进 0/1/2 汉字。
/// 检测对象用 TOC1/2/3 样式，域未展开按 T6 列未判定；条目的页码对齐依赖渲染，不据段落 jc 硬判。
/// </summary>
public sealed class TocFormatCommand
{
    private readonly TocFormatAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public TocFormatCommand(
        TocFormatAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new TocFormatAnalyzer();
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

        if (string.IsNullOrWhiteSpace(path)) { WriteUsage(); return 2; }
        if (!File.Exists(path)) { _error.WriteLine($"找不到文件：{path}"); return 1; }

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

    private void WriteReport(TocFormatResult result)
    {
        _output.WriteLine("=== 目录格式检查报告（R6） ===");
        _output.WriteLine();
        _output.WriteLine($"依据：{result.SourceNote}");
        _output.WriteLine();

        _output.WriteLine(result.TocTitleParagraph > 0
            ? $"识别到目录标题：第 {result.TocTitleParagraph} 段"
            : "未识别到「目录」标题（见下方未检查项）。");
        _output.WriteLine($"识别到的目录条目：{result.Entries.Count} 个，实际检查段落：{result.CheckedParagraphs} 个");
        foreach (var e in result.Entries)
            _output.WriteLine($"  - 第 {e.ParagraphIndex} 段 [{e.Level}] {e.Text}");
        _output.WriteLine();

        _output.WriteLine($"问题：{result.Issues.Count} 条");
        foreach (var issue in result.Issues)
            _output.WriteLine($"  - [{issue.Code}·{issue.RuleId}] {issue.Part}「{Short(issue.Text)}」{issue.Dimension}：期望「{issue.Expected}」，实际「{issue.Actual}」");
        _output.WriteLine();

        _output.WriteLine($"未检查 / 未判定：{result.NotChecked.Count} 项");
        foreach (var item in result.NotChecked)
            _output.WriteLine($"  - {item}");

        if (result.MissingParts.Count > 0)
        {
            _output.WriteLine();
            _output.WriteLine($"缺少部件：{string.Join("、", result.MissingParts)}");
        }
    }

    private static string Short(string text) => text.Length <= 20 ? text : text[..20] + "…";

    private void WriteUsage()
    {
        var usage = new StringBuilder();
        usage.AppendLine("用法：ThesisEngine.Cli analyze-toc <file.docx> [--json]");
        usage.AppendLine("检查目录标题与章/一级/二级目录条目的字体、字号、行距、段前后、缩进（general.yaml R6，draft）。");
        usage.AppendLine("只查不改；按 TOC1/2/3 样式识别条目，TOC 域未展开与页码对齐按未决项列未判定。");
        _output.WriteLine(usage.ToString().TrimEnd());
    }
}
