using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 图题/表题格式检查命令（analyze-captions）。R12，只查不改。
/// 宋体 10.5 磅居中、单倍行距、段前 6 磅；图题段后 12 磅、表题段后 6 磅。
/// 题注相对对象位置、序题间距、R13 表达式编号按未决/渲染依赖列未判定。
/// </summary>
public sealed class CaptionFormatCommand
{
    private readonly CaptionFormatAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public CaptionFormatCommand(
        CaptionFormatAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new CaptionFormatAnalyzer();
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

    private void WriteReport(CaptionFormatResult result)
    {
        _output.WriteLine("=== 图题/表题格式检查报告（R12） ===");
        _output.WriteLine();
        _output.WriteLine($"依据：{result.SourceNote}");
        _output.WriteLine();

        _output.WriteLine($"识别到题注：{result.Captions.Count} 个（图题 {result.FigureCount} / 表题 {result.TableCount}），实际检查：{result.CheckedCaptions} 个");
        foreach (var c in result.Captions)
            _output.WriteLine($"  - 第 {c.ParagraphIndex} 段 [{c.CaptionType}] {c.Text}");
        _output.WriteLine();

        _output.WriteLine($"问题：{result.Issues.Count} 条");
        foreach (var issue in result.Issues)
            _output.WriteLine($"  - [{issue.Code}·{issue.RuleId}] {issue.CaptionType}「{Short(issue.Text)}」{issue.Dimension}：期望「{issue.Expected}」，实际「{issue.Actual}」");
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
        usage.AppendLine("用法：ThesisEngine.Cli analyze-captions <file.docx> [--json]");
        usage.AppendLine("检查图题/表题的字体、字号、居中、单倍行距、段前后（general.yaml R12，draft）。");
        usage.AppendLine("只查不改；按文本锚点识别题注，对象位置/序题间距/R13 表达式编号列未判定。");
        _output.WriteLine(usage.ToString().TrimEnd());
    }
}
