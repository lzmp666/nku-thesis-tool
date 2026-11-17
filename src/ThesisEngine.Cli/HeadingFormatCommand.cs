using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 标题层级格式检查命令（analyze-headings）。R7–R10，只查不改。
/// 章/一级/二级/三级标题的字体、字号、加粗、单倍行距、段前后；对齐与序号间距按条文冲突列未判定。
/// </summary>
public sealed class HeadingFormatCommand
{
    private readonly HeadingFormatAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public HeadingFormatCommand(
        HeadingFormatAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new HeadingFormatAnalyzer();
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

    private void WriteReport(HeadingFormatResult result)
    {
        _output.WriteLine("=== 标题层级格式检查报告（R7–R10） ===");
        _output.WriteLine();
        _output.WriteLine($"依据：{result.SourceNote}");
        _output.WriteLine();

        if (result.ChapterStartParagraph <= 0)
            _output.WriteLine("未识别到章标题，未对任何标题判定格式（见下方未检查项）。");
        else
        {
            var end = result.SectionEndParagraph > 0 ? $"第 {result.SectionEndParagraph} 段（结束标题）之前" : "文档末尾";
            _output.WriteLine($"标题检查区间：第 {result.ChapterStartParagraph} 段（章标题）至{end}");
        }
        _output.WriteLine($"识别到的分级标题：{result.RecognizedHeadings.Count} 个，实际检查：{result.CheckedHeadings} 个");
        foreach (var h in result.RecognizedHeadings)
            _output.WriteLine($"  - 第 {h.ParagraphIndex} 段 [{h.Level} / 大纲{h.OutlineLevel} / {h.NumberingMode}] {h.Text}");
        _output.WriteLine();

        _output.WriteLine($"问题：{result.Issues.Count} 条");
        foreach (var issue in result.Issues)
            _output.WriteLine($"  - [{issue.Code}·{issue.RuleId}] {issue.Level}「{Short(issue.HeadingText)}」{issue.Dimension}：期望「{issue.Expected}」，实际「{issue.Actual}」");
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
        usage.AppendLine("用法：ThesisEngine.Cli analyze-headings <file.docx> [--json]");
        usage.AppendLine("检查章/一级/二级/三级标题的字体、字号、加粗、单倍行距、段前后（general.yaml R7–R10，draft）。");
        usage.AppendLine("只查不改；按大纲级别/标题样式识别标题，对齐与序号间距按条文冲突列未判定。");
        _output.WriteLine(usage.ToString().TrimEnd());
    }
}
