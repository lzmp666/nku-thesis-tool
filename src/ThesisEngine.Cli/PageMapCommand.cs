using System.Text.Encodings.Web;
using System.Text.Json;
using ThesisEngine.Core.Analysis;

namespace ThesisEngine.Cli;

/// <summary>页码估算命令（page-map）：输出每个段落的估算页码，明确标注非 Word 实机页码。</summary>
public sealed class PageMapCommand
{
    private readonly PageNumberEstimator _estimator;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public PageMapCommand(PageNumberEstimator? estimator = null, TextWriter? output = null, TextWriter? error = null)
    {
        _estimator = estimator ?? new PageNumberEstimator();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        string? path = null;
        var json = false;
        var onlyIssues = false;
        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--json": json = true; break;
                case "--issues": onlyIssues = true; break;
                case "--help" or "-h": WriteUsage(); return 0;
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
            _error.WriteLine("缺少 docx 路径。");
            WriteUsage();
            return 2;
        }

        try
        {
            var map = _estimator.Estimate(path!);
            if (json)
            {
                var options = new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, WriteIndented = true };
                _output.WriteLine(JsonSerializer.Serialize(map, options));
            }
            else
            {
                _output.WriteLine($"=== 估算页码映射：{path} ===");
                _output.WriteLine($"总页数（估算）：{map.TotalPages?.ToString() ?? "未判定"}");
                _output.WriteLine(map.Note);
                var rows = onlyIssues
                    ? map.Paragraphs.Where(p => p.StartPage is null || p.ForcedBreakBefore)
                    : map.Paragraphs;
                foreach (var row in rows)
                {
                    var page = row.StartPage is null ? "未判定" : row.EndPage == row.StartPage
                        ? $"第 {row.StartPage} 页" : $"第 {row.StartPage}–{row.EndPage} 页";
                    _output.WriteLine($"  段落 {row.ParagraphIndex}: {page}{(row.ForcedBreakBefore ? "（强制分页）" : "")} | {row.Basis}");
                }
            }
            return 0;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or ArgumentException or UnauthorizedAccessException)
        {
            _error.WriteLine($"页码估算失败：{ex.Message}");
            return 1;
        }
    }

    private void WriteUsage()
    {
        _output.WriteLine("用法: ThesisEngine.Cli page-map <document.docx> [--json] [--issues]");
        _output.WriteLine("  --json    输出机器可读 JSON");
        _output.WriteLine("  --issues  只显示未判定与强制分页的段落");
        _output.WriteLine("注意：页码为 OOXML 静态估算，不是 Word 实机分页，不作为页码合规判据。");
    }
}
