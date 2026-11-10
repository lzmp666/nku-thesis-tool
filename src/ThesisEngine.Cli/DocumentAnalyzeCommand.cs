using System.Xml;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Cli;

/// <summary>
/// 文档级版式检查命令（#112）：纸张、页边距、页眉页脚、页码、目录、空白页。
/// </summary>
public sealed class DocumentAnalyzeCommand
{
    private readonly DocumentAnalyzer _analyzer;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public DocumentAnalyzeCommand(
        DocumentAnalyzer? analyzer = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _analyzer = analyzer ?? new DocumentAnalyzer();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli analyze-document <file.docx>");
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

            _output.WriteLine("=== 文档级版式检查报告 ===");
            _output.WriteLine();
            _output.WriteLine($"文件：{path}");
            _output.WriteLine($"节数：{result.Sections.Count}");
            _output.WriteLine($"奇偶页不同：{(result.EvenAndOddHeaders ? "是" : "否")}");
            _output.WriteLine($"打开时更新域：{(result.UpdateFieldsOnOpen ? "已设置" : "未设置")}");
            _output.WriteLine($"自动目录（TOC 域）：{(result.HasTocField ? "有" : "无")}");
            _output.WriteLine();

            var judged = result.PartReadings.Count(reading => reading.FullyResolved);
            _output.WriteLine($"页眉/页码格式：完整判定 {judged} / {result.PartReadings.Count} 个文字段落" +
                "（未判定的维度逐条列在末尾「本次未检查」里）");
            _output.WriteLine();

            if (result.Sections.Count > 0)
            {
                _output.WriteLine("--- 页面设置读数 ---");
                foreach (var section in result.Sections)
                {
                    WriteSection(section);
                }
                _output.WriteLine();
            }

            if (result.PartReadings.Count > 0)
            {
                _output.WriteLine("--- 页眉 / 页码格式读数 ---");
                foreach (var reading in result.PartReadings)
                {
                    WritePartReading(reading);
                }
                _output.WriteLine();
            }

            var clauseIssues = result.ClauseIssues.ToList();
            var conventionIssues = result.ConventionIssues.ToList();

            if (clauseIssues.Count > 0)
            {
                _output.WriteLine($"⚠️ 不符合规范：{clauseIssues.Count} 处");
                foreach (var issue in clauseIssues)
                {
                    WriteIssue(issue);
                }
                _output.WriteLine();
            }

            if (conventionIssues.Count > 0)
            {
                _output.WriteLine($"排版惯例提示：{conventionIssues.Count} 处（规范未规定，不计入规范符合性）");
                foreach (var issue in conventionIssues)
                {
                    WriteIssue(issue);
                }
                _output.WriteLine();
            }

            if (clauseIssues.Count == 0 && conventionIssues.Count == 0)
            {
                var partial = result.PartReadings.Count(reading => !reading.FullyResolved);
                _output.WriteLine(partial > 0
                    ? $"已判定的项目未发现问题；但有 {partial} 个页眉/页码段落的部分维度未判定，" +
                        "不能据此说文档合规（见下方「本次未检查」）。"
                    : "已检查的文档级项目均符合。");
                _output.WriteLine();
            }

            _output.WriteLine("--- 本次未检查 ---");
            foreach (var item in result.UncheckedItems)
            {
                _output.WriteLine($"  - {item}");
            }

            return 0;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or
            UnauthorizedAccessException or XmlException or ArgumentException)
        {
            _error.WriteLine($"检查失败：{ex.Message}");
            return 1;
        }
    }

    private void WriteSection(SectionInfo section)
    {
        _output.WriteLine($"第 {section.Index} 节：{section.PageWidthCm:0.##} × {section.PageHeightCm:0.##} cm");
        _output.WriteLine($"  页边距：上 {section.TopCm:0.##} / 下 {section.BottomCm:0.##} / " +
            $"左 {section.LeftCm:0.##} / 右 {section.RightCm:0.##} cm");
        _output.WriteLine($"  页眉距边界 {section.HeaderCm:0.##} cm，页脚距边界 {section.FooterCm:0.##} cm，" +
            $"装订线 {section.GutterCm:0.##} cm");

        var format = section.PageNumberFormat switch
        {
            // 首节没设 pgNumType 时是 Word 默认的十进制，说成「继承上一节」会误导。
            null when section.Index == 1 => "未设置（Word 默认阿拉伯数字）",
            null => "未设置（继承上一节）",
            "decimal" => "阿拉伯数字",
            "upperRoman" => "大写罗马数字",
            "lowerRoman" => "小写罗马数字",
            var other => other,
        };
        var start = section.PageNumberStart?.ToString() ?? "未设置";
        _output.WriteLine($"  页码：格式 {format}，起始 {start}");

        var header = section.HeaderCount == 0
            ? "无页眉引用"
            : section.HasHeaderContent ? "有内容" : "引用了但为空";
        var footer = section.FooterCount == 0
            ? "无页脚引用"
            : section.HasFooterContent ? "有内容" : "引用了但为空";
        _output.WriteLine($"  页眉：{header}（引用 {section.HeaderCount} 个），" +
            $"页脚：{footer}（引用 {section.FooterCount} 个）");
        if (section.TitlePage)
        {
            _output.WriteLine("  首页不同：已设置");
        }
    }

    private void WritePartReading(PartFormatReading reading)
    {
        var font = reading.FontResolved
            ? $"中文字体 {reading.EastAsiaFont ?? "未设置"} / 西文字体 {reading.WesternFont ?? "未设置"}"
            : "字体：样式链里没读到，未判定";
        var size = reading.SizeResolved
            ? $"{reading.SizeHalfPoints / 2.0:0.#} 磅"
            : "未判定";
        var alignment = reading.AlignmentResolved ? reading.Alignment : "未判定";

        _output.WriteLine($"  {reading.PartName} 第 {reading.ParagraphIndex} 段" +
            $"（{(reading.WesternOnly ? "全西文" : "含中文")}，" +
            $"样式链{(reading.StylesPartAvailable ? "可用" : "不可用")}）「{Short(reading.Text)}」");
        _output.WriteLine($"    {font}｜字号 {size}｜对齐 {alignment}" +
            $"｜{(reading.FullyResolved ? "完整判定" : "有维度未判定，见文末")}");
    }

    private static string Short(string text) => text.Length <= 20 ? text : text[..20] + "…";

    private void WriteIssue(DocumentIssue issue)
    {
        _output.WriteLine($"  [{issue.Code}] {issue.Title}");
        _output.WriteLine(issue.SectionIndex > 0
            ? $"    位置：第 {issue.SectionIndex} 节"
            : "    位置：全文");
        _output.WriteLine($"    当前：{issue.Current}");
        _output.WriteLine($"    应为：{issue.Expected}");
        _output.WriteLine($"    依据：{issue.Basis}");
    }
}
