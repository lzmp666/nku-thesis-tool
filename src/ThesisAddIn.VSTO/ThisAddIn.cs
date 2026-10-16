using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;
using ThesisAddIn.Protocol;
using ThesisAddIn.VSTO.Engine;
using ThesisAddIn.VSTO.Ui;

namespace ThesisAddIn.VSTO
{
    /// <summary>
    /// Word 侧入口：创建中文问题侧边栏，把分析、定位、修复和批注锚点联动接到一起。
    ///
    /// 重要：Word 会锁住正在编辑的 .docx，引擎无法直接读取，因此所有分析/修复
    /// 都基于 Document.SaveCopyAs 生成的临时副本；修复结果另存到用户指定的新文件，
    /// 原始文件与当前打开中的文档都不会被覆盖。
    /// </summary>
    public partial class ThisAddIn
    {
        internal const string TaskPaneTitle = "论文格式检查";

        /// <summary>批注正文里用于回填规则编号的片段，例如“规则：demo/east-asian-font。”。</summary>
        private static readonly Regex RuleIdPattern = new Regex(@"规则[:：]\s*(?<id>[A-Za-z0-9_\-/]+)", RegexOptions.Compiled);

        private CustomTaskPane _taskPane;
        private DiagnosticsPane _pane;
        private EngineGateway _gateway;
        private string _gatewayError;

        public bool IsEngineConfigured
        {
            get { return _gateway != null; }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _gateway = EngineGateway.Create(out _gatewayError);

            _pane = new DiagnosticsPane();
            _pane.AnalyzeRequested += OnAnalyzeRequested;
            _pane.FixCurrentRequested += OnFixCurrentRequested;
            _pane.FixAllRequested += OnFixAllRequested;
            _pane.DiagnosticSelected += OnDiagnosticSelected;

            _taskPane = CustomTaskPanes.Add(_pane, TaskPaneTitle);
            _taskPane.Visible = true;

            Application.WindowSelectionChange += OnWindowSelectionChange;

            _pane.SetStatus(string.IsNullOrEmpty(_gatewayError)
                ? "准备就绪。点击“分析当前文档”开始检查。"
                : _gatewayError);

            WriteStartupLog("插件启动完成；引擎已配置：" + (IsEngineConfigured ? "是" : "否")
                + "；侧边栏：" + (_taskPane != null ? "已创建" : "创建失败"));
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                Application.WindowSelectionChange -= OnWindowSelectionChange;
            }
            catch (Exception)
            {
                // Word 正在退出时事件可能已失效，忽略。
            }

            _gateway = null;
        }

        /// <summary>由 ThisAddIn.Designer.cs 生成的初始化流程调用，绑定启动/关闭事件。</summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        private void OnAnalyzeRequested(object sender, EventArgs e)
        {
            if (!EnsureGateway())
            {
                return;
            }

            string failure;
            string workCopy = CreateWorkingCopy(out failure);
            if (workCopy == null)
            {
                _pane.ShowFailure(failure);
                return;
            }

            _pane.SetBusy(true, "正在分析文档…（基于当前内容的临时副本，不修改原文件）");
            try
            {
                ShowAnalysisResult(_gateway.Analyze(workCopy), "分析完成：");
            }
            finally
            {
                _pane.SetBusy(false, null);
                TryDelete(workCopy);
            }
        }

        private void OnDiagnosticSelected(object sender, EventArgs e)
        {
            EngineDiagnostic diagnostic = _pane.SelectedDiagnostic;
            if (diagnostic == null)
            {
                return;
            }

            string status = NavigateTo(diagnostic);
            if (!string.IsNullOrEmpty(status))
            {
                _pane.SetStatus(status);
            }
        }

        private void OnFixCurrentRequested(object sender, EventArgs e)
        {
            EngineDiagnostic diagnostic = _pane.SelectedDiagnostic;
            if (diagnostic == null)
            {
                _pane.SetStatus("请先在列表中选择一条问题，再执行“修复此项”。");
                return;
            }

            if (!string.Equals(diagnostic.FixLevel, "auto", StringComparison.OrdinalIgnoreCase))
            {
                _pane.SetStatus("该问题不能自动修复（处理方式：" + DiagnosticsPane.DescribeFixLevel(diagnostic.FixLevel)
                    + "）。请按说明手动调整，或先与规则负责人确认后再纳入自动修复。");
                return;
            }

            RunRepair(diagnostic, "当前选中项 " + diagnostic.RuleId);
        }

        private void OnFixAllRequested(object sender, EventArgs e)
        {
            RunRepair(null, "全部可自动修复项");
        }

        /// <summary>
        /// 修复闭环：临时副本 → 生成计划 → 中文预览确认 → 选择输出文件 → 执行 → 重新分析。
        /// 任何一步失败都停止，不会产生半成品。
        /// </summary>
        private void RunRepair(EngineDiagnostic diagnostic, string scopeLabel)
        {
            if (!EnsureGateway())
            {
                return;
            }

            string failure;
            string workCopy = CreateWorkingCopy(out failure);
            if (workCopy == null)
            {
                _pane.ShowFailure(failure);
                return;
            }

            try
            {
                _pane.SetBusy(true, "正在生成修复计划…");
                EngineResponse planResponse = _gateway.Plan(workCopy, diagnostic);
                if (planResponse == null || !planResponse.Ok || planResponse.Plan == null)
                {
                    _pane.ShowFailure(EngineGateway.DescribeError(planResponse != null ? planResponse.Error : null));
                    return;
                }

                EnginePatchPlan plan = planResponse.Plan;
                if (plan.Patches.Count == 0)
                {
                    _pane.SetStatus("没有可自动修复的项目（" + scopeLabel + "）。这些项目可能需要人工确认，或者已经修复。");
                    return;
                }

                using (var preview = new FixPreviewDialog(plan, scopeLabel))
                {
                    if (preview.ShowDialog() != DialogResult.OK)
                    {
                        _pane.SetStatus("已取消修复，文档没有任何改动。");
                        return;
                    }
                }

                string outputPath = AskOutputPath();
                if (string.IsNullOrEmpty(outputPath))
                {
                    _pane.SetStatus("未选择输出文件，已取消修复。");
                    return;
                }

                _pane.SetStatus("正在写入新文件…（原文件保持不变）");
                EngineResponse applyResponse = _gateway.Apply(workCopy, outputPath, diagnostic);
                if (applyResponse == null || !applyResponse.Ok || applyResponse.Apply == null)
                {
                    _pane.ShowFailure(EngineGateway.DescribeError(applyResponse != null ? applyResponse.Error : null));
                    return;
                }

                ReportApplyResult(applyResponse.Apply);
                ShowAnalysisResult(_gateway.Analyze(applyResponse.Apply.OutputPath), "已按修复结果重新分析：");
            }
            catch (Exception exception)
            {
                _pane.ShowFailure("修复过程出现异常，已停止：" + exception.Message);
            }
            finally
            {
                _pane.SetBusy(false, null);
                TryDelete(workCopy);
            }
        }

        private void ReportApplyResult(EngineApplyResult result)
        {
            var message = new System.Text.StringBuilder();
            message.Append("修复完成：已应用 ").Append(result.Applied.Count).Append(" 处");
            if (result.Failures.Count > 0)
            {
                message.Append("；失败 ").Append(result.Failures.Count).Append(" 处（未产生任何部分写入）");
            }

            message.Append("。输出文件：").Append(result.OutputPath);
            foreach (EnginePatchFailure failure in result.Failures)
            {
                message.Append(Environment.NewLine).Append("· 失败：").Append(failure.Patch.RuleId)
                    .Append(" @ ").Append(failure.Patch.Location).Append(" — ").Append(failure.Reason);
            }

            _pane.SetStatus(message.ToString());
        }

        private void ShowAnalysisResult(EngineResponse response, string prefix)
        {
            if (response != null && response.Ok && response.Result != null)
            {
                _pane.ShowAnalysis(response.Result);
                _pane.SetStatus(prefix + "当前剩余 " + response.Result.Diagnostics.Count + " 条问题（含未自动修复项）。");
                return;
            }

            _pane.ShowFailure(EngineGateway.DescribeError(response != null ? response.Error : null));
        }

        /// <summary>把一条诊断定位到原文：选中对应段落（选中即为可见高亮，不修改文档内容）。</summary>
        private string NavigateTo(EngineDiagnostic diagnostic)
        {
            DocumentLocation location;
            string parseError;
            if (!DocumentLocator.TryParse(diagnostic.Location, out location, out parseError))
            {
                return parseError;
            }

            Word.Document document = TryGetActiveDocument();
            Word.Range range;
            string rangeError;
            if (!DocumentLocator.TryGetParagraphRange(document, location, out range, out rangeError))
            {
                return rangeError;
            }

            try
            {
                range.Select();
                return "已定位到" + DocumentLocator.Describe(location) + "：" + diagnostic.RuleId + "。";
            }
            catch (Exception exception)
            {
                return "定位失败：" + exception.Message;
            }
        }

        /// <summary>
        /// 批注锚点到建议卡片的联动。Word 对象模型没有“鼠标悬停”事件，
        /// 因此这里监听选择变化（点击或选中锚点）作为回退实现；
        /// 该限制已在 docs 中如实记录，不宣称支持纯悬停。
        /// </summary>
        private void OnWindowSelectionChange(object selection)
        {
            Word.Range range = selection as Word.Range;
            if (_pane == null || range == null)
            {
                return;
            }

            string ruleId = TryReadRuleIdFromComments(range);
            if (string.IsNullOrEmpty(ruleId))
            {
                return;
            }

            if (_pane.SelectByRuleId(ruleId))
            {
                _pane.SetStatus("已按批注锚点激活建议卡片：" + ruleId + "（Word 无悬停事件，此处为点击/选中激活）");
            }
        }

        private static string TryReadRuleIdFromComments(Word.Range selection)
        {
            try
            {
                Word.Range paragraphRange = selection.Paragraphs.Count > 0 ? selection.Paragraphs[1].Range : selection;
                Word.Comments comments = paragraphRange.Comments;
                for (int i = 1; i <= comments.Count; i++)
                {
                    string text = comments[i].Range.Text;
                    Match match = RuleIdPattern.Match(text ?? string.Empty);
                    if (match.Success)
                    {
                        return match.Groups["id"].Value;
                    }
                }
            }
            catch (Exception)
            {
                // 选区可能已被关闭，忽略本次联动。
            }

            return null;
        }

        /// <summary>
        /// 生成供引擎读取的临时副本。Word 锁定打开中的 .docx，引擎无法直接打开该文件；
        /// SaveCopyAs 生成的副本同时包含尚未保存的修改。
        /// </summary>
        private string CreateWorkingCopy(out string error)
        {
            error = null;
            Word.Document document = TryGetActiveDocument();
            if (document == null)
            {
                error = "当前没有打开的文档，请先打开一个 .docx 文件。";
                return null;
            }

            try
            {
                string directory = Path.Combine(Path.GetTempPath(), "NkuThesisTool");
                Directory.CreateDirectory(directory);
                string path = Path.Combine(directory, "work-" + Guid.NewGuid().ToString("N") + ".docx");
                document.SaveCopyAs(path);
                return path;
            }
            catch (Exception exception)
            {
                error = "无法生成用于检查的临时副本：" + exception.Message;
                return null;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
                // 临时文件残留不影响使用。
            }
        }

        private static string AskOutputPath()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "选择修复结果的保存位置（不会覆盖原文件）";
                dialog.Filter = "Word 文档 (*.docx)|*.docx";
                dialog.FileName = "论文格式检查-已修复.docx";
                dialog.OverwritePrompt = true;
                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
            }
        }

        private bool EnsureGateway()
        {
            if (_gateway != null)
            {
                return true;
            }

            _pane.ShowFailure(_gatewayError ?? "引擎未配置。");
            return false;
        }

        private Word.Document TryGetActiveDocument()
        {
            try
            {
                return Application.ActiveDocument;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 把启动结果写到 %LOCALAPPDATA%\NkuThesisTool\addin.log。
        /// 侧边栏或引擎加载失败时，用户可以直接把这个文件发给维护者。
        /// </summary>
        internal static void WriteStartupLog(string message)
        {
            try
            {
                string directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "NkuThesisTool");
                Directory.CreateDirectory(directory);
                File.AppendAllText(
                    Path.Combine(directory, "addin.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + message + Environment.NewLine);
            }
            catch (Exception)
            {
                // 记录失败不影响插件运行。
            }
        }
    }
}
