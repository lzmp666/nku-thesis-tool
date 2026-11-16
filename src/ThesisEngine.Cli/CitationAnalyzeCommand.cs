using ThesisEngine.Core.Citation;

namespace ThesisEngine.Cli;

public sealed class CitationAnalyzeCommand
{
    public int Execute(string[] arguments, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);
        if (arguments.Length != 3 || !string.Equals(arguments[0], "analyze", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(arguments[1], "--citations", StringComparison.Ordinal) || string.IsNullOrWhiteSpace(arguments[2]))
        {
            error.WriteLine("用法：ThesisEngine.Cli analyze --citations <file.docx>");
            return 2;
        }

        var inputPath = arguments[2];
        try
        {
            var fullPath = Path.GetFullPath(inputPath);
            if (!string.Equals(Path.GetExtension(fullPath), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                error.WriteLine("输入格式不受支持：仅支持 .docx 文件。");
                return 1;
            }
            if (!File.Exists(fullPath))
            {
                error.WriteLine($"找不到输入文件：{inputPath}");
                return 1;
            }

            var result = new CitationConsistencyAnalyzer().Analyze(fullPath);
            WriteReport(result, output, fullPath);
            return 0;
        }
        catch (UnauthorizedAccessException)
        {
            error.WriteLine("分析失败：无法访问输入文件，请检查文件权限。");
            return 1;
        }
        catch (IOException)
        {
            error.WriteLine("分析失败：无法读取文档，请检查文件是否损坏或路径是否有效。");
            return 1;
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidDataException or NotSupportedException or System.Xml.XmlException)
        {
            error.WriteLine($"分析失败：文档、路径或输入格式无效。{exception.Message}");
            return 1;
        }
    }

    private static void WriteReport(CitationAnalysisResult result, TextWriter output, string path)
    {
        output.WriteLine("=== 正文引用与参考文献一致性检查 ===");
        output.WriteLine("模式：顺序编码制（只检查，不修改文档）");
        output.WriteLine($"输入文件：{path}");
        output.WriteLine($"引用标注：{result.CitationCount}；参考文献条目：{result.ReferenceCount}");
        output.WriteLine($"问题数量：{result.Findings.Count}");
        output.WriteLine("规则版本：citation-0.1.0；检查：编号连续、一一对应、标点位置、条目字段、同文献复用编号");

        output.WriteLine("问题详情：");
        if (result.Findings.Count == 0)
        {
            output.WriteLine("  （无）");
        }
        else
        {
            for (var index = 0; index < result.Findings.Count; index++)
            {
                var finding = result.Findings[index];
                output.WriteLine($"  [{index + 1}] {finding.RuleId}");
                output.WriteLine($"      位置：word/document.xml#paragraph[{finding.ParagraphIndex}]");
                output.WriteLine($"      实际：{finding.Actual}");
                output.WriteLine($"      说明：{finding.Message}");
            }
        }

        if (result.CitationFieldCount > 0)
            output.WriteLine($"引用管理器域：识别到 {result.CitationFieldCount} 个域；仅统计，不解析 Zotero/EndNote 域语义。");
        foreach (var scope in result.UnsupportedScopes)
            output.WriteLine($"未检查：{scope}");
        output.WriteLine("提示：只报告可识别的问题，不会改写正文或参考文献；规则未覆盖的项目不作判断。");
    }
}
