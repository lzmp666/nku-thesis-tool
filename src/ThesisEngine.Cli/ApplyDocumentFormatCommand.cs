using ThesisEngine.Core;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Cli;

/// <summary>
/// 文档级格式写入（#112 第二档）：只改 sectPr 下的纸张、页边距、页眉页脚距、装订线。
/// 写完立刻复诊，把「修了什么」「剩下什么」「为什么没修」分开写清楚。
/// </summary>
public sealed class ApplyDocumentFormatCommand
{
    private static readonly string[] AutoFixCodes = { "D1", "D2", "D3", "D4" };

    private readonly DocumentFormatter _formatter;
    private readonly DocumentAnalyzer _analyzer;
    private readonly UndoManager _undo;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public ApplyDocumentFormatCommand(
        DocumentFormatter? formatter = null,
        DocumentAnalyzer? analyzer = null,
        UndoManager? undo = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _formatter = formatter ?? new DocumentFormatter();
        _analyzer = analyzer ?? new DocumentAnalyzer();
        _undo = undo ?? new UndoManager();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli apply-document-format <input.docx> [output.docx]");
            return 2;
        }

        var inputPath = args[0];
        if (!File.Exists(inputPath))
        {
            _error.WriteLine($"找不到文件：{inputPath}");
            return 1;
        }

        var outputPath = args.Length >= 2
            ? args[1]
            : Path.Combine(
                Path.GetDirectoryName(inputPath) ?? "",
                Path.GetFileNameWithoutExtension(inputPath) + "_文档格式.docx");

        try
        {
            _output.WriteLine("=== 文档级格式写入（只动 sectPr）===");
            _output.WriteLine();
            _output.WriteLine($"输入：{inputPath}");

            var before = _analyzer.Analyze(inputPath);

            var backupId = _undo.BackupBeforeModify(
                inputPath,
                outputPath,
                "文档级格式：纸张 A4、四边页边距、页眉页脚距、装订线");
            _output.WriteLine($"备份 ID：{backupId}（不满意可运行 undo 撤销）");
            _output.WriteLine();

            var result = _formatter.Apply(inputPath, outputPath);

            _output.WriteLine($"扫描 {result.SectionCount} 节，已修改 {result.Fixes.Count} 处属性");
            _output.WriteLine();

            if (result.Fixes.Count > 0)
            {
                _output.WriteLine("--- 已修改 ---");
                foreach (var fix in result.Fixes)
                {
                    _output.WriteLine($"  第 {fix.SectionIndex} 节  [{fix.Code}] {fix.Description}");
                }
                _output.WriteLine();
            }
            else
            {
                _output.WriteLine("纸张与页边距已经符合规范，没有需要修改的属性。");
                _output.WriteLine();
            }

            if (result.Skips.Count > 0)
            {
                _output.WriteLine($"--- 没有自动改的 {result.Skips.Count} 处 ---");
                foreach (var skip in result.Skips)
                {
                    _output.WriteLine($"  - {skip}");
                }
                _output.WriteLine();
            }

            WriteFollowUp(before, outputPath);
            WriteUncheckedItems();

            _output.WriteLine($"输出：{outputPath}");
            _output.WriteLine("提示：用 Word 打开确认；不满意运行 undo 撤销到修改前状态。");
            return 0;
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException or ArgumentException)
        {
            _error.WriteLine($"写入失败：{ex.Message}");
            return 1;
        }
    }

    /// <summary>复诊：D1–D4 必须清零；其它项本来就不自动改，逐条列出来别让人以为修完了。</summary>
    private void WriteFollowUp(DocumentAnalysisResult before, string outputPath)
    {
        var after = _analyzer.Analyze(outputPath);

        var beforeAuto = before.Issues.Count(issue => AutoFixCodes.Contains(issue.Code));
        var afterAuto = after.Issues.Count(issue => AutoFixCodes.Contains(issue.Code));
        var beforeOther = before.Issues.Count - beforeAuto;
        var afterOther = after.Issues.Count - afterAuto;

        _output.WriteLine("--- 复诊 ---");
        _output.WriteLine($"可自动修的 D1–D4：修改前 {beforeAuto} 处 → 修改后 {afterAuto} 处");
        _output.WriteLine($"只报告不改的项：修改前 {beforeOther} 处 → 修改后 {afterOther} 处");

        if (afterAuto > 0)
        {
            _output.WriteLine("  ⚠️ D1–D4 还有剩余，可能是 sectPr 缺少 pgSz / pgMar 元素（见上方「没有自动改」）：");
            foreach (var issue in after.Issues.Where(issue => AutoFixCodes.Contains(issue.Code)))
            {
                _output.WriteLine($"    [{issue.Code}] {issue.Title}｜第 {issue.SectionIndex} 节｜当前：{issue.Current}");
            }
        }

        var remaining = after.Issues.Where(issue => !AutoFixCodes.Contains(issue.Code)).ToList();
        if (remaining.Count > 0)
        {
            _output.WriteLine($"  剩余 {remaining.Count} 处需人工处理：");
            foreach (var issue in remaining)
            {
                _output.WriteLine($"    [{issue.Code}] {issue.Title}｜当前：{issue.Current}");
            }
        }

        _output.WriteLine("  注意：复诊 0 只说明 D1–D4 这几项没发现差异，不等于整篇文档合规；" +
            "未检查项见 analyze-document 报告末尾。");
        _output.WriteLine();
    }

    private void WriteUncheckedItems()
    {
        _output.WriteLine("--- 本次未处理（issue 明确要求只报告不改）---");
        _output.WriteLine("  - 目录：手打目录要换成 TOC 域，自动改会动正文内容");
        _output.WriteLine("  - 页码：PAGE 域的位置、编号格式、是否续前节，动它会改域结构");
        _output.WriteLine("  - 空白页：删段落属于删内容，误删的风险担不起，只给操作步骤");
        _output.WriteLine("  - 页眉页脚字体字号与居中：写在 header/footer 部件里，不在 sectPr，超出允许范围");
        _output.WriteLine("  - 奇偶页不同、首页不同、打开时更新域：写在 settings.xml / sectPr 其它元素里，不碰");
        _output.WriteLine("  - 分节数量与分节类型：改了会重排全文");
        _output.WriteLine();
    }
}
