using System;
using System.Collections.Generic;
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
        private static readonly Regex ExpectedValuePattern = new Regex(@"期望值[:：]\s*(?<value>.*?)(?:[；;。\r\n]|$)", RegexOptions.Compiled);
        private static readonly Regex ActualValuePattern = new Regex(@"实际值[:：]\s*(?<value>.*?)(?:[；;。\r\n]|$)", RegexOptions.Compiled);

        private CustomTaskPane _taskPane;
        private DiagnosticsPane _pane;
        private EngineGateway _gateway;
        private string _gatewayError;
        private string _workingCopySource = "已保存的文件内容";
        private System.Windows.Forms.Timer _documentHoverTimer;
        private int _lastHoveredWindowHandle = -1;
        private int _lastHoveredRangeStart = -1;
        private int _lastHoveredRangeEnd = -1;
        private string _lastHoveredCommentKey;

        public bool IsEngineConfigured
        {
            get { return _gateway != null; }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 任何未捕获的异常都写进日志：未激活的 Office 环境下部分 Word 调用会抛
            // 含糊的 COM 错误，只有日志能定位到具体位置。
            AppDomain.CurrentDomain.UnhandledException += (s, args) =>
                WriteStartupLog("未处理异常：" + args.ExceptionObject);
            System.Windows.Forms.Application.ThreadException += (s, args) =>
                WriteStartupLog("界面线程异常：" + args.Exception);

            _gateway = EngineGateway.Create(out _gatewayError);

            Application.WindowActivate += OnWindowActivate;
            Application.WindowSelectionChange += OnWindowSelectionChange;

            // Word 没有文档正文的鼠标悬停事件。用低频 UI 计时器查询鼠标下的
            // Word.Range，再通过批注锚点切换建议卡片；COM 调用仍在 Word UI 线程执行。
            _documentHoverTimer = new System.Windows.Forms.Timer { Interval = 250 };
            _documentHoverTimer.Tick += OnDocumentHoverTimerTick;
            _documentHoverTimer.Start();

            // Word can start on its backstage page with no document window. In that
            // state, reading Application.ActiveWindow may throw a COM exception and
            // make Office disable the add-in (LoadBehavior 3 -> 2) before the user
            // opens a document. WindowActivate will create the pane when a window
            // becomes available.
            try
            {
                Word.Window activeWindow = Application.ActiveWindow;
                if (activeWindow != null)
                {
                    EnsureTaskPaneForWindow(activeWindow);
                }
            }
            catch (Exception exception)
            {
                WriteStartupLog("启动时当前没有可用的 Word 文档窗口；等待窗口激活后创建侧边栏："
                    + exception.Message);
            }

            WriteStartupLog("插件启动完成；引擎已配置：" + (IsEngineConfigured ? "是" : "否")
                + "；侧边栏：" + (_taskPane != null ? "已创建" : "创建失败"));
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (_documentHoverTimer != null)
                {
                    _documentHoverTimer.Stop();
                    _documentHoverTimer.Tick -= OnDocumentHoverTimerTick;
                    _documentHoverTimer.Dispose();
                    _documentHoverTimer = null;
                }
                Application.WindowActivate -= OnWindowActivate;
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
            ActivatePaneFromSender(sender);
            string workCopy = null;
            try
            {
                WriteStartupLog("收到“分析当前文档”点击。");
                if (!EnsureGateway())
                {
                    WriteStartupLog("分析中止：引擎未配置。");
                    return;
                }

                string failure;
                workCopy = CreateWorkingCopy(out failure);
                if (workCopy == null)
                {
                    WriteStartupLog("分析中止：临时副本失败 - " + failure);
                    _pane.ShowFailure(failure);
                    return;
                }

                WriteStartupLog("临时副本已生成，开始调用引擎。");
                _pane.SetBusy(true, "正在分析文档…（基于" + _workingCopySource + "的临时副本，不修改原文件）");
                ShowAnalysisResult(_gateway.Analyze(workCopy), "分析完成：");
                WriteStartupLog("分析调用返回。");
            }
            catch (Exception exception)
            {
                WriteStartupLog("分析出现异常：" + exception);
                _pane.ShowFailure("分析出现异常，已停止：" + exception.Message);
            }
            finally
            {
                _pane.SetBusy(false, null);
                TryDelete(workCopy);
            }
        }

        private void OnDiagnosticSelected(object sender, EventArgs e)
        {
            ActivatePaneFromSender(sender);
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
            ActivatePaneFromSender(sender);
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
            ActivatePaneFromSender(sender);
            WriteStartupLog("收到“预览并修复全部”点击。");
            RunRepair(null, "全部可自动修复项");
        }

        /// <summary>
        /// Word 每个文档窗口都需要自己的任务窗格和诊断列表。启动时只创建首个窗口的
        /// 窗格会导致后来打开的文档没有侧栏；WindowActivate 同时覆盖新窗口和窗口切换。
        /// </summary>
        private void OnWindowActivate(Word.Document document, Word.Window window)
        {
            EnsureTaskPaneForWindow(window);
        }

        private void EnsureTaskPaneForWindow(Word.Window window)
        {
            if (window == null)
            {
                _taskPane = null;
                _pane = null;
                return;
            }

            int windowHandle;
            try
            {
                windowHandle = window.Hwnd;
            }
            catch (Exception exception)
            {
                WriteStartupLog("无法读取 Word 窗口标识，未创建侧边栏：" + exception.Message);
                return;
            }

            for (int i = 0; i < CustomTaskPanes.Count; i++)
            {
                CustomTaskPane existingTaskPane = CustomTaskPanes[i];
                try
                {
                    Word.Window paneWindow = existingTaskPane.Window as Word.Window;
                    if (paneWindow == null)
                    {
                        continue;
                    }

                    if (paneWindow.Hwnd == windowHandle)
                    {
                        _taskPane = existingTaskPane;
                        _pane = existingTaskPane.Control as DiagnosticsPane;
                        return;
                    }
                }
                catch (Exception)
                {
                    // 任务窗格可能正随文档窗口关闭；跳过失效的 COM 窗口引用。
                }
            }

            DiagnosticsPane pane = CreateDiagnosticsPane();
            CustomTaskPane taskPane = null;
            try
            {
                taskPane = CustomTaskPanes.Add(pane, TaskPaneTitle, window);
                taskPane.Visible = true;
                _taskPane = taskPane;
                _pane = pane;
                pane.SetStatus(string.IsNullOrEmpty(_gatewayError)
                    ? "准备就绪。点击“分析当前文档”开始检查。"
                    : _gatewayError);
                WriteStartupLog("已为 Word 窗口 " + windowHandle + " 创建中文侧边栏。");
            }
            catch (Exception exception)
            {
                if (taskPane != null)
                {
                    try
                    {
                        CustomTaskPanes.Remove(taskPane);
                    }
                    catch (Exception)
                    {
                        // Word 关闭窗口期间集合可能已不可访问；继续清理控件自身。
                    }
                }

                pane.Dispose();
                _taskPane = null;
                _pane = null;
                WriteStartupLog("创建 Word 任务窗格失败：" + exception);
            }
        }

        private DiagnosticsPane CreateDiagnosticsPane()
        {
            var pane = new DiagnosticsPane();
            pane.AnalyzeRequested += OnAnalyzeRequested;
            pane.FixCurrentRequested += OnFixCurrentRequested;
            pane.FixAllRequested += OnFixAllRequested;
            pane.DiagnosticSelected += OnDiagnosticSelected;
            return pane;
        }

        private void ActivatePaneFromSender(object sender)
        {
            DiagnosticsPane pane = sender as DiagnosticsPane;
            if (pane == null)
            {
                return;
            }

            _pane = pane;
            _taskPane = null;
            for (int i = 0; i < CustomTaskPanes.Count; i++)
            {
                CustomTaskPane taskPane = CustomTaskPanes[i];
                if (object.ReferenceEquals(taskPane.Control, pane))
                {
                    _taskPane = taskPane;
                    break;
                }
            }
        }

        /// <summary>
        /// 修复闭环：临时副本 → 生成计划 → 中文预览确认 → 选择输出文件 → 执行 → 重新分析。
        /// 任何一步失败都停止，不会产生半成品。
        /// </summary>
        private void RunRepair(EngineDiagnostic diagnostic, string scopeLabel)
        {
            WriteStartupLog("收到修复请求：" + scopeLabel + "。");
            if (!EnsureGateway())
            {
                WriteStartupLog("修复中止：引擎未配置。");
                return;
            }

            string failure;
            string workCopy = CreateWorkingCopy(out failure);
            if (workCopy == null)
            {
                WriteStartupLog("修复中止：临时副本失败 - " + failure);
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

                WriteStartupLog("修复计划已生成：待修改 " + plan.Patches.Count + " 处，保留 " + plan.Deferred.Count + " 处。");
                using (var preview = new FixPreviewDialog(plan, scopeLabel))
                {
                    WriteStartupLog("准备显示修复预览窗口。");
                    if (preview.ShowDialog() != DialogResult.OK)
                    {
                        WriteStartupLog("用户在预览窗口取消。");
                        _pane.SetStatus("已取消修复，文档没有任何改动。");
                        return;
                    }

                    WriteStartupLog("用户确认预览。");
                }

                string outputPath = AskOutputPath();
                if (string.IsNullOrEmpty(outputPath))
                {
                    WriteStartupLog("用户未选择输出文件。");
                    _pane.SetStatus("未选择输出文件，已取消修复。");
                    return;
                }

                WriteStartupLog("开始写入：" + outputPath);
                _pane.SetStatus("正在写入新文件…（原文件保持不变）");
                EngineResponse applyResponse = _gateway.Apply(workCopy, outputPath, diagnostic);
                if (applyResponse == null || !applyResponse.Ok || applyResponse.Apply == null)
                {
                    string applyError = EngineGateway.DescribeError(applyResponse != null ? applyResponse.Error : null);
                    WriteStartupLog("写入失败：" + applyError);
                    _pane.ShowFailure(applyError);
                    return;
                }

                EngineApplyResult applied = applyResponse.Apply;
                if (applied.Failures != null && applied.Failures.Count > 0)
                {
                    string failureSummary = ReportApplyResult(applied);
                    WriteStartupLog(failureSummary);
                    _pane.ShowFailure(failureSummary + Environment.NewLine + "原始文档未修改，也没有生成部分修复文件。");
                    return;
                }

                if (string.IsNullOrWhiteSpace(applied.OutputPath) || !File.Exists(applied.OutputPath))
                {
                    const string missingOutput = "修复引擎未生成可复查的输出文件；原始文档保持不变。";
                    WriteStartupLog(missingOutput);
                    _pane.ShowFailure(missingOutput);
                    return;
                }

                string applySummary = ReportApplyResult(applied);
                WriteStartupLog(applySummary);

                // 输出文件是实际修复结果；原文仍保持打开且未修改。复查输出文件，
                // 同时保留侧边栏与当前打开的原文对应，避免把输出问题定位到原文。
                EngineResponse outputCheck = _gateway.Analyze(applied.OutputPath);
                if (outputCheck != null && outputCheck.Ok && outputCheck.Result != null)
                {
                    EngineResponse sourceCheck = _gateway.Analyze(workCopy);
                    if (sourceCheck != null && sourceCheck.Ok && sourceCheck.Result != null)
                    {
                        _pane.ShowAnalysis(sourceCheck.Result);
                    }

                    _pane.SetStatus(applySummary + Environment.NewLine
                        + "修复结果复查：输出文件剩余 " + outputCheck.Result.Diagnostics.Count + " 条问题。"
                        + Environment.NewLine + "当前列表仍对应未修改的原始文档；修复结果已另存到上方路径。");
                }
                else
                {
                    _pane.SetStatus(applySummary + Environment.NewLine
                        + "输出文件已生成，但复查未完成："
                        + EngineGateway.DescribeError(outputCheck != null ? outputCheck.Error : null));
                }
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

        private string ReportApplyResult(EngineApplyResult result)
        {
            var message = new System.Text.StringBuilder();
            message.Append(result.Failures != null && result.Failures.Count > 0
                ? "修复未完成：已应用 "
                : "修复完成：已应用 ").Append(result.Applied.Count).Append(" 处");
            if (result.Applied.Count > 0)
            {
                message.Append("（");
                for (int i = 0; i < result.Applied.Count; i++)
                {
                    if (i > 0) message.Append("、");
                    message.Append(result.Applied[i].RuleId);
                }
                message.Append("）");
            }

            if (result.Failures.Count > 0)
            {
                message.Append("；失败 ").Append(result.Failures.Count).Append(" 处（未产生任何部分写入）");
            }

            if (string.IsNullOrWhiteSpace(result.OutputPath))
            {
                message.Append("。未生成输出文件。");
            }
            else
            {
                message.Append("。输出文件：").Append(result.OutputPath);
            }
            foreach (EnginePatchFailure failure in result.Failures)
            {
                message.Append(Environment.NewLine).Append("· 失败：").Append(failure.Patch.RuleId)
                    .Append(" @ ").Append(failure.Patch.Location).Append(" — ").Append(failure.Reason);
            }

            return message.ToString();
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
        /// 批注锚点到建议卡片的选择联动。
        /// </summary>
        private void OnWindowSelectionChange(object selection)
        {
            try
            {
                EnsureTaskPaneForWindow(Application.ActiveWindow);
            }
            catch (Exception)
            {
                // Word 关闭文档或切换窗口期间，ActiveWindow 可能暂时不可用。
            }

            // Word raises WindowSelectionChange with a Word.Selection COM object.
            // It is not a Word.Range, so read its Range explicitly.
            Word.Selection wordSelection = selection as Word.Selection;
            Word.Range range = wordSelection != null ? wordSelection.Range : selection as Word.Range;
            if (_pane == null || range == null)
            {
                return;
            }

            string ruleId;
            string expected;
            string actual;
            if (!TryReadIssueFromComments(range, out ruleId, out expected, out actual))
            {
                return;
            }

            if (_pane.SelectByComment(ruleId, expected, actual, navigateToSource: false))
            {
                _pane.SetStatus("已按批注锚点激活建议卡片：" + ruleId + "。");
            }
        }

        /// <summary>
        /// Word 没有文档正文的 MouseHover 事件。定时查询鼠标屏幕坐标下的 Range，
        /// 并仅在命中范围变化时检查批注锚点，避免每个计时周期遍历批注。
        /// </summary>
        private void OnDocumentHoverTimerTick(object sender, EventArgs e)
        {
            try
            {
                Word.Window window = Application.ActiveWindow;
                if (window == null)
                {
                    ResetHoveredRange();
                    return;
                }

                EnsureTaskPaneForWindow(window);
                if (_pane == null)
                {
                    ResetHoveredRange();
                    return;
                }

                System.Drawing.Point cursor = Cursor.Position;
                Word.Range range = window.RangeFromPoint(cursor.X, cursor.Y) as Word.Range;
                if (range == null)
                {
                    ResetHoveredRange();
                    return;
                }

                int windowHandle = window.Hwnd;
                int start = range.Start;
                int end = range.End;
                if (windowHandle != _lastHoveredWindowHandle)
                {
                    _lastHoveredCommentKey = null;
                }
                if (windowHandle == _lastHoveredWindowHandle
                    && start == _lastHoveredRangeStart
                    && end == _lastHoveredRangeEnd)
                {
                    return;
                }

                _lastHoveredWindowHandle = windowHandle;
                _lastHoveredRangeStart = start;
                _lastHoveredRangeEnd = end;

                string ruleId;
                string expected;
                string actual;
                if (TryReadIssueFromComments(range, out ruleId, out expected, out actual))
                {
                    string commentKey = ruleId + "\u001f" + expected + "\u001f" + actual;
                    if (!string.Equals(commentKey, _lastHoveredCommentKey, StringComparison.Ordinal)
                        && _pane.SelectByComment(ruleId, expected, actual, navigateToSource: false))
                    {
                        _lastHoveredCommentKey = commentKey;
                    }
                }
                else
                {
                    _lastHoveredCommentKey = null;
                }
            }
            catch (Exception)
            {
                // 切换窗口、切换到 Backstage 或 Word 正在关闭时，坐标查询可能失败。
                // 悬停是辅助联动，失败时保留现有卡片，不打断分析和编辑。
                ResetHoveredRange();
            }
        }

        private void ResetHoveredRange()
        {
            _lastHoveredWindowHandle = -1;
            _lastHoveredRangeStart = -1;
            _lastHoveredRangeEnd = -1;
            _lastHoveredCommentKey = null;
        }

        /// <summary>
        /// 只读取与当前选区锚点相交的批注，避免同一段有多条批注时总是激活第一条。
        /// 若选区覆盖多条批注或无法唯一匹配，则不切换侧边栏。
        /// </summary>
        private static bool TryReadIssueFromComments(Word.Range selection, out string ruleId, out string expected, out string actual)
        {
            ruleId = null;
            expected = null;
            actual = null;
            try
            {
                Word.Comments comments = selection.Document.Comments;
                var matches = new List<string[]>();
                for (int i = 1; i <= comments.Count; i++)
                {
                    Word.Comment comment = comments[i];
                    Word.Range scope = comment.Scope;
                    bool intersects = selection.Start == selection.End
                        ? selection.Start >= scope.Start && selection.Start <= scope.End
                        : selection.Start < scope.End && scope.Start < selection.End;
                    if (!intersects)
                    {
                        continue;
                    }

                    string text = comment.Range.Text ?? string.Empty;
                    Match ruleMatch = RuleIdPattern.Match(text);
                    Match expectedMatch = ExpectedValuePattern.Match(text);
                    Match actualMatch = ActualValuePattern.Match(text);
                    if (ruleMatch.Success && expectedMatch.Success && actualMatch.Success)
                    {
                        matches.Add(new[]
                        {
                            ruleMatch.Groups["id"].Value.Trim(),
                            expectedMatch.Groups["value"].Value.Trim(),
                            actualMatch.Groups["value"].Value.Trim()
                        });
                    }
                }

                if (matches.Count != 1)
                {
                    return false;
                }

                ruleId = matches[0][0];
                expected = matches[0][1];
                actual = matches[0][2];
                return true;
            }
            catch (Exception)
            {
                // 选区或批注可能已被关闭，忽略本次联动。
                return false;
            }
        }

        /// <summary>
        /// 生成供引擎读取的临时副本。
        ///
        /// 优先使用 Word 自己的 SaveCopyAs：它包含尚未保存的修改。部分环境会拒绝该命令
        /// （本机实测：Office 未激活时返回 0x800A1704），此时退回文件系统复制，检查的
        /// 是磁盘上已保存的内容——这一点会如实反映在面板状态文字里。
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

            string directory = Path.Combine(Path.GetTempPath(), "NkuThesisTool");
            string path = Path.Combine(directory, "work-" + Guid.NewGuid().ToString("N") + ".docx");
            try
            {
                Directory.CreateDirectory(directory);
            }
            catch (Exception exception)
            {
                error = "无法创建临时目录：" + exception.Message;
                return null;
            }

            try
            {
                document.SaveCopyAs(path);
                if (File.Exists(path))
                {
                    _workingCopySource = "当前内容";
                    return path;
                }

                WriteStartupLog("SaveCopyAs 未生成文件，改用文件复制。");
            }
            catch (Exception exception)
            {
                WriteStartupLog("SaveCopyAs 不可用（" + exception.Message + "），改用文件复制。");
            }

            try
            {
                string source = document.FullName;
                if (string.IsNullOrEmpty(source) || !File.Exists(source))
                {
                    error = "当前文档还没有保存到磁盘，无法生成用于检查的副本。请先保存文档。";
                    return null;
                }

                File.Copy(source, path, true);
                _workingCopySource = "已保存的文件内容";
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
                // 带时间戳的默认文件名：修复引擎拒绝写已存在的文件，
                // 固定名字会让第二次修复被安全规则挡下。
                dialog.FileName = "论文格式检查-已修复-" + DateTime.Now.ToString("MMdd-HHmmss") + ".docx";
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

