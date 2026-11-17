using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 正文段落格式检查命令（analyze-body-text）。R11 补齐，只查不改。
/// 只判两端对齐 / 首行缩进 / 固定行距 / 段前后；字体字号由 analyze-styles 的 #43 规则负责。
/// </summary>
public sealed class BodyTextFormatCommand
{
    private readonly BodyTextFormatAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public BodyTextFormatCommand(
        BodyTextFormatAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new BodyTextFormatAnalyzer();
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

    private void WriteReport(BodyTextFormatResult result)
    {
        _output.WriteLine("=== 正文段落格式检查报告（R11 补齐） ===");
        _output.WriteLine();
        _output.WriteLine($"依据：{result.SourceNote}");
        _output.WriteLine();

        if (result.ChapterStartParagraph <= 0 || result.SectionEndParagraph <= 0)
            _output.WriteLine("正文边界未完整识别，未对任何段落判定排版格式（见下方未检查项）。");
        else
            _output.WriteLine($"正文区间：第 {result.ChapterStartParagraph} 段（章标题）至第 {result.SectionEndParagraph} 段（结束标题）之前");
        _output.WriteLine($"实际检查正文段落：{result.CheckedParagraphs} 段");
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
        }
    }

    private void WriteUsage()
    {
        var usage = new StringBuilder();
        usage.AppendLine("用法：ThesisEngine.Cli analyze-body-text <file.docx> [--json]");
        usage.AppendLine("检查正文段落的两端对齐、首行缩进 2 汉字符、固定值行距 20 磅、段前后 0 磅（general.yaml R11，draft）。");
        usage.AppendLine("只查不改；字体与字号由 #43 的正文规则负责，本命令不重复；判不出的维度列未判定。");
        _output.WriteLine(usage.ToString().TrimEnd());
    }
}
