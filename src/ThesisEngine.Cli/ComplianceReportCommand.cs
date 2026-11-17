using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 统一合规报告命令（report）。一次调用跑完全部「只查不改」分析器，
/// 汇成一份报告；单个分析器失败不影响其它，失败会明确写进报告。
/// </summary>
public sealed class ComplianceReportCommand
{
    private readonly ComplianceReportService _service;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public ComplianceReportCommand(
        ComplianceReportService? service = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _service = service ?? new ComplianceReportService();
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
            var report = _service.Analyze(path);
            if (json)
            {
                _output.WriteLine(JsonSerializer.Serialize(report, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                }));
            }
            else
            {
                WriteReport(report);
            }
            return 0;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException or ArgumentException)
        {
            _error.WriteLine($"生成报告失败：{ex.Message}");
            return 1;
        }
    }

    private void WriteReport(ComplianceReport report)
    {
        _output.WriteLine("=== 论文格式统一合规报告 ===");
        _output.WriteLine($"文件：{report.InputPath}");
        _output.WriteLine("依据：南开大学研究生学位论文格式规范 general.yaml（条款状态 draft，需人工确认）；只查不改。");
        _output.WriteLine();

        _output.WriteLine("检查器执行情况：");
        foreach (var section in report.Sections)
        {
            var status = section.Succeeded
                ? $"完成（规范 {section.IssueCount} / 惯例 {section.ConventionCount} / 未判定 {section.NotCheckedCount}）"
                : $"失败：{section.Error}";
            _output.WriteLine($"  - {section.Name} [{section.Analyzer}]：{status}");
        }
        _output.WriteLine();

        _output.WriteLine("未运行成功的检查器不代表合规，相关维度本次没有结论。");
        _output.WriteLine();

        _output.WriteLine($"规范问题：{report.TotalIssues} 条");
        foreach (var issue in report.Issues)
        {
            var location = issue.ParagraphIndex.HasValue ? $"第 {issue.ParagraphIndex} 段" : "全局";
            _output.WriteLine($"  - [{issue.Analyzer}/{issue.Code}] {issue.Title}（{location}）：{issue.Detail}");
        }
        _output.WriteLine();

        _output.WriteLine($"排版惯例（不计入规范符合性）：{report.TotalConventions} 条");
        foreach (var issue in report.Conventions)
            _output.WriteLine($"  - [{issue.Analyzer}/{issue.Code}] {issue.Title}：{issue.Detail}");
        _output.WriteLine();

        _output.WriteLine($"未检查 / 未判定：{report.TotalNotChecked} 项");
        foreach (var item in report.NotChecked)
            _output.WriteLine($"  - {item}");
        _output.WriteLine();

        _output.WriteLine("说明：本报告为静态 XML 检查，不替代 Word 实机渲染核验；规则状态 draft。" +
                          "「未判定」不是「合规」，「缺部件」也不是「合规」。");
    }

    private void WriteUsage()
    {
        var usage = new StringBuilder();
        usage.AppendLine("用法：ThesisEngine.Cli report <file.docx> [--json]");
        usage.AppendLine("一次运行全部只查不改的分析器（页面/结构、样式、图表、公式、脚注、引文、摘要、尾注），");
        usage.AppendLine("汇成一份统一报告。规范问题、排版惯例、未判定三者分列；单个检查器失败会写进报告，");
        usage.AppendLine("其余照常产出，失败的维度不据此判定合规。");
        _output.WriteLine(usage.ToString().TrimEnd());
    }
}
