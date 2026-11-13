using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// 图表相关的段落遍历与题注识别入口。
///
/// 之前三处逻辑各自用 <c>body.Elements(w:p)</c>，只取正文直接子段落，
/// 表格单元格里的题注（例如“续表”落在表内第一行）会被漏掉。这里统一走
/// <see cref="Of"/>，正文段落与表格内段落都包含在内。
/// </summary>
public static class FigureTableParagraphs
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    /// <summary>按文档顺序返回正文段落与表格内段落。</summary>
    public static IEnumerable<XElement> Of(XElement body)
    {
        foreach (var element in body.Elements())
        {
            if (element.Name == W + "p")
            {
                yield return element;
            }
            else if (element.Name == W + "tbl")
            {
                foreach (var paragraph in element.Descendants(W + "p")) yield return paragraph;
            }
        }
    }

    /// <summary>
    /// 题注识别：支持“图 1-1 / 表 2-3 / 续表 4.1 / Figure 1-1 / Table 2-3”，
    /// 编号分隔符接受 - . — – ． 等常见写法。识别不出时返回 <see cref="CaptionKind"/> 之外的 null。
    /// </summary>
    public static CaptionKind? Classify(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        var match = CaptionPattern.Match(text.Trim());
        if (!match.Success) return null;

        var label = match.Groups["label"].Value;
        if (label.Contains("表", StringComparison.Ordinal) ||
            label.StartsWith("Table", StringComparison.OrdinalIgnoreCase))
        {
            return CaptionKind.Table;
        }

        return CaptionKind.Figure;
    }

    /// <summary>是否为“续表”题注（续表在表格内部，单独处理）。</summary>
    public static bool IsContinuedTable(string? text) =>
        !string.IsNullOrWhiteSpace(text) && text.TrimStart().StartsWith("续表", StringComparison.Ordinal);

    public static string Text(XElement paragraph) =>
        string.Concat(paragraph.Descendants(W + "t").Select(t => t.Value)).Trim();

    private static readonly Regex CaptionPattern = new(
        @"^(?<label>续?表|图|Table|Figure|Fig\.)\s*(?<number>(?:[A-Z]|[0-9０-９]+)(?:[.\-—–．][0-9０-９]+)*)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
}
