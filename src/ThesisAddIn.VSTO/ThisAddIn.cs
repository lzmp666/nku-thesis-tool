using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
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
    /// 都基于 Word 当前内存内容生成的临时副本；修复结果另存到用户指定的新文件，
    /// 原始文件与当前打开中的文档都不会被覆盖。
    /// </summary>
    public partial class ThisAddIn
    {
        internal const string TaskPaneTitle = "论文格式检查";

        /// <summary>批注正文里用于回填规则编号的片段，例如“规则：demo/east-asian-font。”。</summary>
        private static readonly Regex RuleIdPattern = new Regex(@"规则[:：]\s*(?<id>[A-Za-z0-9_\-/]+)", RegexOptions.Compiled);
        private static readonly Regex ExpectedValuePattern = new Regex(@"期望值[:：]\s*(?<value>.*?)(?:[；;。\r\n]|$)", RegexOptions.Compiled);
        private static readonly Regex ActualValuePattern = new Regex(@"实际值[:：]\s*(?<value>.*?)(?:[；;。\r\n]|$)", RegexOptions.Compiled);
        private static readonly Regex CommentLocationPattern = new Regex(
            @"\[NKU-LOCATION:(?<location>word/document\.xml#paragraph\[\d+\]/run\[\d+\])\]",
            RegexOptions.Compiled);

        private CustomTaskPane _taskPane;
        private DiagnosticsPane _pane;
        private EngineGateway _gateway;
        private string _gatewayError;
        private System.Windows.Forms.Timer _documentHoverTimer;
        private CancellationTokenSource _operationCancellation;
        private int _lastHoveredWindowHandle = -1;
        private int _lastHoveredCursorX = -1;
        private int _lastHoveredCursorY = -1;
        private string _lastHoveredCommentKey;
        private bool _isNavigatingToDiagnostic;

        private sealed class DocumentOperationContext
        {
            public Word.Document Document { get; set; }
            public Word.Window Window { get; set; }
            public int WindowHandle { get; set; }
            public string DocumentName { get; set; }
            public string DocumentFullName { get; set; }
            public string SnapshotPath { get; set; }
            public string Fingerprint { get; set; }
        }

        public bool IsEngineConfigured
        {
            get { return _gateway != null; }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ThesisRibbon(OpenDiagnosticsPane);
        }

        private void OpenDiagnosticsPane()
        {
            try
            {
                EnsureTaskPaneForWindow(Application.ActiveWindow);
                if (_taskPane == null)
                {
                    throw new InvalidOperationException("当前 Word 窗口没有可用的检查面板。");
                }

                // WindowActivate 会复用已有窗格；用户关闭的窗格仍在集合中，
                // 因此必须明确恢复 Visible，不能仅再次调用创建方法。
                _taskPane.Visible = true;
                WriteStartupLog("已通过功能区打开当前文档的检查面板。");
            }
            catch (Exception exception)
            {
                WriteStartupLog("通过功能区打开检查面板失败：" + exception);
                MessageBox.Show("未能打开检查面板。请先打开 Word 文档，再重试。",
                    TaskPaneTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
            _operationCancellation = new CancellationTokenSource();

            Application.WindowActivate += OnWindowActivate;
            Application.WindowSelectionChange += OnWindowSelectionChange;

            // Word 没有文档正文的鼠标悬停事件。用 UI 计时器查询鼠标下的
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
            if (_operationCancellation != null)
            {
                _operationCancellation.Cancel();
            }

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
            if (_operationCancellation != null)
            {
                _operationCancellation.Dispose();
                _operationCancellation = null;
            }
        }

        /// <summary>由 ThisAddIn.Designer.cs 生成的初始化流程调用，绑定启动/关闭事件。</summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        private async void OnAnalyzeRequested(object sender, EventArgs e)
        {
            ActivatePaneFromSender(sender);
            DiagnosticsPane pane = sender as DiagnosticsPane ?? _pane;
            string workCopy = null;
            string annotatedPath = null;
            try
            {
                WriteStartupLog("收到“分析当前文档”点击。");
                if (!EnsureGateway(pane))
                {
                    WriteStartupLog("分析中止：引擎未配置。");
                    return;
                }
                EngineGateway gateway = _gateway;
                CancellationToken cancellationToken = GetOperationCancellationToken();

                DocumentOperationContext operation;
                string failure;
                if (!TryCreateWorkingCopy(pane, out operation, out failure))
                {
                    WriteStartupLog("分析中止：临时副本失败 - " + failure);
                    ShowFailure(pane, failure);
                    return;
                }
                workCopy = operation.SnapshotPath;
                annotatedPath = CreateAnnotatedOutputPath(operation);

                WriteStartupLog("已为 Word 窗口 " + operation.WindowHandle + " 生成当前内容快照，开始分析并创建标注副本。");
                SetBusy(pane, true, "正在分析并生成标注副本…（读取当前 Word 内容，原文不变）");
                EngineResponse response = await gateway.AnnotateAsync(workCopy, annotatedPath, cancellationToken);
                if (response == null || !response.Ok || response.Result == null || response.Annotation == null
                    || !PathsEqual(response.Annotation.OutputPath, annotatedPath)
                    || !File.Exists(annotatedPath))
                {
                    string detail = response != null && response.Ok && response.Annotation != null
                        && !PathsEqual(response.Annotation.OutputPath, annotatedPath)
                        ? "标注引擎返回了与请求不一致的输出路径，已停止打开该文件。"
                        : EngineGateway.DescribeError(response != null ? response.Error : null);
                    ShowFailure(pane, detail);
                    return;
                }

                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    TryDelete(annotatedPath);
                    annotatedPath = null;
                    ShowFailure(pane, failure);
                    return;
                }

                string openError;
                Word.Document annotatedDocument = OpenAnnotatedCopy(annotatedPath, out openError);
                if (annotatedDocument == null)
                {
                    ShowFailure(pane, "批注副本已生成，但 Word 未能打开它：" + openError
                        + Environment.NewLine + "文件位置：" + annotatedPath);
                    return;
                }

                DiagnosticsPane annotatedPane = GetPaneForWindow(annotatedDocument.ActiveWindow);
                ShowAnalysis(annotatedPane, response.Result);
                SetStatus(annotatedPane, "分析完成：已生成并打开标注副本，共 "
                    + response.Result.Diagnostics.Count + " 条问题；原文保持不变。\n" + annotatedPath);
                SetStatus(pane, "分析完成：已在独立窗口打开标注副本，原文保持不变。\n" + annotatedPath);
                WriteStartupLog("已打开标注副本，批注数 " + response.Annotation.CommentCount + "，窗口 "
                    + SafeWindowHandle(annotatedDocument.ActiveWindow) + "。");
            }
            catch (Exception exception)
            {
                WriteStartupLog("分析出现异常：" + exception);
                ShowFailure(pane, "分析出现异常，已停止：" + exception.Message);
            }
            finally
            {
                SetBusy(pane, false, null);
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

            string status = NavigateTo(_pane, diagnostic);
            if (!string.IsNullOrEmpty(status))
            {
                _pane.SetStatus(status);
            }
        }

        private async void OnFixCurrentRequested(object sender, EventArgs e)
        {
            ActivatePaneFromSender(sender);
            DiagnosticsPane pane = _pane;
            EngineDiagnostic diagnostic = pane != null ? pane.SelectedDiagnostic : null;
            if (diagnostic == null)
            {
                SetStatus(pane, "请先悬停原文标注，或用上一项/下一项选择问题，再执行“修复此项”。");
                return;
            }

            if (!string.Equals(diagnostic.FixLevel, "auto", StringComparison.OrdinalIgnoreCase))
            {
                SetStatus(pane, "该问题不能自动修复（处理方式：" + DiagnosticsPane.DescribeFixLevel(diagnostic.FixLevel)
                    + "）。请按说明手动调整，或先与规则负责人确认后再纳入自动修复。");
                return;
            }

            await RunRepair(pane, diagnostic, "当前选中项 " + diagnostic.RuleId);
        }

        private async void OnFixAllRequested(object sender, EventArgs e)
        {
            ActivatePaneFromSender(sender);
            WriteStartupLog("收到“预览并修复全部”点击。");
            await RunRepair(_pane, null, "全部可自动修复项");
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
                // Keep the Chinese diagnostic cards readable and make all three actions
                // visible without relying on a very wide monitor or a second task pane.
                taskPane.Width = 520;
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

        private DiagnosticsPane GetPaneForWindow(Word.Window window)
        {
            EnsureTaskPaneForWindow(window);
            if (window == null)
            {
                return null;
            }

            try
            {
                int windowHandle = window.Hwnd;
                for (int i = 0; i < CustomTaskPanes.Count; i++)
                {
                    CustomTaskPane taskPane = CustomTaskPanes[i];
                    Word.Window paneWindow = taskPane.Window as Word.Window;
                    if (paneWindow != null && paneWindow.Hwnd == windowHandle)
                    {
                        taskPane.Visible = true;
                        return taskPane.Control as DiagnosticsPane;
                    }
                }
            }
            catch (Exception exception)
            {
                WriteStartupLog("无法取得标注副本的侧边栏：" + exception.Message);
            }

            return null;
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
        private async Task RunRepair(DiagnosticsPane pane, EngineDiagnostic diagnostic, string scopeLabel)
        {
            string workCopy = null;
            string applyOutput = null;
            string finalOutput = null;
            bool keepApplyOutput = false;
            try
            {
                WriteStartupLog("收到修复请求：" + scopeLabel + "。");
                if (!EnsureGateway(pane))
                {
                    WriteStartupLog("修复中止：引擎未配置。");
                    return;
                }
                EngineGateway gateway = _gateway;
                CancellationToken cancellationToken = GetOperationCancellationToken();

                DocumentOperationContext operation;
                string failure;
                if (!TryCreateWorkingCopy(pane, out operation, out failure))
                {
                    WriteStartupLog("修复中止：临时副本失败 - " + failure);
                    ShowFailure(pane, failure);
                    return;
                }
                workCopy = operation.SnapshotPath;

                SetBusy(pane, true, "正在生成修复计划…");
                EngineResponse planResponse = await gateway.PlanAsync(workCopy, diagnostic, cancellationToken);
                if (planResponse == null || !planResponse.Ok || planResponse.Plan == null)
                {
                    ShowFailure(pane, EngineGateway.DescribeError(planResponse != null ? planResponse.Error : null));
                    return;
                }

                EnginePatchPlan plan = planResponse.Plan;
                if (plan.Patches.Count == 0)
                {
                    SetStatus(pane, "没有可自动修复的项目（" + scopeLabel + "）。这些项目可能需要人工确认，或者已经修复。");
                    return;
                }

                WriteStartupLog("修复计划已生成：待修改 " + plan.Patches.Count + " 处，保留 " + plan.Deferred.Count + " 处。");
                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    ShowFailure(pane, failure);
                    return;
                }

                DialogResult previewResult = pane.InvokeOnUiThread(() =>
                {
                    using (var preview = new FixPreviewDialog(plan, scopeLabel))
                    {
                        WriteStartupLog("准备显示修复预览窗口。");
                        return preview.ShowDialog();
                    }
                });
                if (previewResult != DialogResult.OK)
                {
                    WriteStartupLog("用户在预览窗口取消。");
                    SetStatus(pane, "已取消修复，文档没有任何改动。");
                    return;
                }

                WriteStartupLog("用户确认预览。");

                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    ShowFailure(pane, failure);
                    return;
                }

                finalOutput = pane.InvokeOnUiThread(() => AskOutputPath());
                if (string.IsNullOrEmpty(finalOutput))
                {
                    WriteStartupLog("用户未选择输出文件。");
                    SetStatus(pane, "未选择输出文件，已取消修复。");
                    return;
                }

                if (File.Exists(finalOutput))
                {
                    SetStatus(pane, "为避免覆盖已有文件，请重新选择一个尚不存在的输出文件名。");
                    finalOutput = null;
                    return;
                }

                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    ShowFailure(pane, failure);
                    return;
                }

                applyOutput = CreateTemporaryDocxPath("repair");

                /*
                 * The dialogs above are synchronous UI-thread calls. After they return,
                 * the engine can continue on a worker thread; DiagnosticsPane marshals
                 * subsequent status and result updates back to the same UI thread.
                 */
                WriteStartupLog("开始生成修复中间副本。");
                SetStatus(pane, "正在写入修复副本…（原文保持不变）");
                EngineResponse applyResponse = await gateway.ApplyAsync(workCopy, applyOutput, diagnostic, cancellationToken);
                if (applyResponse == null || !applyResponse.Ok || applyResponse.Apply == null)
                {
                    string applyError = EngineGateway.DescribeError(applyResponse != null ? applyResponse.Error : null);
                    WriteStartupLog("写入失败：" + applyError);
                    ShowFailure(pane, applyError);
                    return;
                }

                EngineApplyResult applied = applyResponse.Apply;
                if (applied.Failures != null && applied.Failures.Count > 0)
                {
                    string failureSummary = ReportApplyResult(applied);
                    WriteStartupLog(failureSummary);
                    ShowFailure(pane, failureSummary + Environment.NewLine + "原始文档未修改，也没有生成部分修复文件。");
                    return;
                }

                if (!PathsEqual(applied.OutputPath, applyOutput) || !File.Exists(applyOutput))
                {
                    const string invalidOutput = "修复引擎返回的输出路径与本次请求不一致，或未生成修复文件；原始文档保持不变。";
                    WriteStartupLog(invalidOutput);
                    ShowFailure(pane, invalidOutput);
                    return;
                }

                string applySummary = ReportApplyResult(applied);
                WriteStartupLog(applySummary);

                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    ShowFailure(pane, failure);
                    return;
                }

                SetStatus(pane, "正在重新分析修复结果并刷新工具批注…");
                EngineResponse refreshed = await gateway.AnnotateAsync(applied.OutputPath, finalOutput, cancellationToken);
                if (refreshed == null || !refreshed.Ok || refreshed.Result == null || refreshed.Annotation == null
                    || !PathsEqual(refreshed.Annotation.OutputPath, finalOutput)
                    || !File.Exists(finalOutput))
                {
                    keepApplyOutput = File.Exists(applyOutput);
                    string refreshError = EngineGateway.DescribeError(refreshed != null ? refreshed.Error : null);
                    if (refreshed != null && refreshed.Ok && refreshed.Annotation != null
                        && !PathsEqual(refreshed.Annotation.OutputPath, finalOutput))
                    {
                        refreshError = "标注引擎返回了与请求不一致的输出路径。";
                    }
                    ShowFailure(pane, applySummary + Environment.NewLine + "修复已写入中间副本，但刷新批注失败："
                        + refreshError + (keepApplyOutput ? Environment.NewLine + "可恢复的修复文件：" + applyOutput : string.Empty));
                    return;
                }

                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    TryDelete(finalOutput);
                    finalOutput = null;
                    ShowFailure(pane, failure);
                    return;
                }

                string openError;
                Word.Document resultDocument = OpenAnnotatedCopy(finalOutput, out openError);
                if (resultDocument == null)
                {
                    ShowFailure(pane, applySummary + Environment.NewLine + "修复结果已生成，但 Word 未能打开它："
                        + openError + Environment.NewLine + "文件位置：" + finalOutput);
                    return;
                }

                DiagnosticsPane resultPane = GetPaneForWindow(resultDocument.ActiveWindow);
                ShowAnalysis(resultPane, refreshed.Result);
                SetStatus(resultPane, applySummary + Environment.NewLine
                    + "修复结果已重新分析，剩余 " + refreshed.Result.Diagnostics.Count + " 条问题；批注已刷新。\n"
                    + "原文保持不变。结果文件：" + finalOutput);
                SetStatus(pane, "修复完成：已在独立窗口打开并复查结果副本，当前文档保持不变。\n" + finalOutput);
                WriteStartupLog("已打开修复并刷新批注的结果副本；剩余问题 "
                    + refreshed.Result.Diagnostics.Count + " 条。");
            }
            catch (Exception exception)
            {
                ShowFailure(pane, "修复过程出现异常，已停止：" + exception.Message);
            }
            finally
            {
                SetBusy(pane, false, null);
                TryDelete(workCopy);
                if (!keepApplyOutput)
                {
                    TryDelete(applyOutput);
                }
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
                message.Append("。修复副本已生成。");
            }
            foreach (EnginePatchFailure failure in result.Failures)
            {
                message.Append(Environment.NewLine).Append("· 失败：").Append(failure.Patch.RuleId)
                    .Append(" @ ").Append(failure.Patch.Location).Append(" — ").Append(failure.Reason);
            }

            return message.ToString();
        }

        /// <summary>把一条诊断定位到原文：选中对应段落（选中即为可见高亮，不修改文档内容）。</summary>
        private string NavigateTo(DiagnosticsPane pane, EngineDiagnostic diagnostic)
        {
            DocumentLocation location;
            string parseError;
            if (!DocumentLocator.TryParse(diagnostic.Location, out location, out parseError))
            {
                return parseError;
            }

            Word.Document document;
            string contextError;
            if (!TryGetDocumentForPane(pane, out document, out _, out contextError))
            {
                return contextError;
            }

            Word.Range range;
            string rangeError;
            if (!DocumentLocator.TryGetParagraphRange(document, location, out range, out rangeError))
            {
                return rangeError;
            }

            try
            {
                // 程序化选中段落会触发 WindowSelectionChange。重叠批注无法由段落选区
                // 唯一辨认，忽略这一次回调，避免导航到后续问题时又切回第一条批注。
                _isNavigatingToDiagnostic = true;
                range.Select();
                return "已定位到" + DocumentLocator.Describe(location) + "：" + diagnostic.RuleId + "。";
            }
            catch (Exception exception)
            {
                return "定位失败：" + exception.Message;
            }
            finally
            {
                _isNavigatingToDiagnostic = false;
            }
        }

        /// <summary>
        /// 批注锚点到建议卡片的选择联动。
        /// </summary>
        private void OnWindowSelectionChange(object selection)
        {
            if (_isNavigatingToDiagnostic)
            {
                return;
            }

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
            if (_pane == null || _pane.IsBusy || range == null)
            {
                return;
            }

            string ruleId;
            string expected;
            string actual;
            string location;
            int matchCount;
            if (!TryReadIssueFromComments(range, out ruleId, out expected, out actual, out location, out matchCount))
            {
                return;
            }

            if (_pane.SelectByComment(ruleId, expected, actual, location, navigateToSource: false))
            {
                WriteStartupLog("选择联动：story=" + range.StoryType + "; start=" + range.Start
                    + "; end=" + range.End + "; location=" + location + "; matches=" + matchCount);
                _pane.SetStatus(matchCount > 1
                    ? "此处有多条建议，当前关联第一条；可用上一项/下一项浏览。"
                    : "当前批注已关联到问题：" + ruleId + "。");
            }
        }

        /// <summary>
        /// Word 没有文档正文的 MouseHover 事件。鼠标坐标变化时查询对应范围并检查锚点：
        /// RangeFromPoint 有时会对同一段中的不同位置返回相同范围，因此不能只比较 Word.Range。
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
                if (_pane == null || _pane.IsBusy)
                {
                    ResetHoveredRange();
                    return;
                }

                int windowHandle = window.Hwnd;
                if (windowHandle != _lastHoveredWindowHandle)
                {
                    _lastHoveredCommentKey = null;
                }
                System.Drawing.Point cursor = Cursor.Position;
                if (windowHandle == _lastHoveredWindowHandle
                    && cursor.X == _lastHoveredCursorX
                    && cursor.Y == _lastHoveredCursorY)
                {
                    return;
                }

                _lastHoveredWindowHandle = windowHandle;
                _lastHoveredCursorX = cursor.X;
                _lastHoveredCursorY = cursor.Y;

                Word.Range range = window.RangeFromPoint(cursor.X, cursor.Y) as Word.Range;
                // Hover sync is for the document body only. Word can return an
                // obscured comment-story range under our task-pane buttons.
                // Native comment clicks are handled by WindowSelectionChange.
                if (range == null || range.StoryType != Word.WdStoryType.wdMainTextStory)
                {
                    ResetHoveredRange();
                    return;
                }

                // RangeFromPoint can return a nearby range even over a task pane.
                // Do not let moving to navigation/repair controls replace the target.
                Word.Range hitBounds = range.Duplicate;
                if (hitBounds.Start == hitBounds.End)
                {
                    hitBounds.MoveEnd(Word.WdUnits.wdCharacter, 1);
                }
                int rangeLeft, rangeTop, rangeWidth, rangeHeight;
                window.GetPoint(out rangeLeft, out rangeTop, out rangeWidth, out rangeHeight, hitBounds);
                if (cursor.X < rangeLeft || cursor.X >= rangeLeft + Math.Max(1, rangeWidth)
                    || cursor.Y < rangeTop || cursor.Y >= rangeTop + Math.Max(1, rangeHeight))
                {
                    ResetHoveredRange();
                    return;
                }

                string ruleId;
                string expected;
                string actual;
                string location;
                int matchCount;
                if (TryReadIssueFromComments(range, out ruleId, out expected, out actual, out location, out matchCount))
                {
                    string commentKey = location + "\u001f" + ruleId + "\u001f" + expected + "\u001f" + actual;
                    if (!string.Equals(commentKey, _lastHoveredCommentKey, StringComparison.Ordinal)
                        && _pane.SelectByComment(ruleId, expected, actual, location, navigateToSource: false))
                    {
                        _lastHoveredCommentKey = commentKey;
                        WriteStartupLog("悬停联动：story=" + range.StoryType + "; start=" + range.Start
                            + "; end=" + range.End + "; location=" + location);
                        if (matchCount > 1)
                        {
                            _pane.SetStatus("此处有多条建议，已显示第一条；可用上一项/下一项浏览。");
                        }
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
            _lastHoveredCursorX = -1;
            _lastHoveredCursorY = -1;
            _lastHoveredCommentKey = null;
        }

        /// <summary>
        /// 只读取与当前选区锚点相交的批注，避免同一段有多条批注时总是激活第一条。
        /// 若选区覆盖多条批注或无法唯一匹配，则不切换侧边栏。
        /// </summary>
        private static bool TryReadIssueFromComments(
            Word.Range selection,
            out string ruleId,
            out string expected,
            out string actual,
            out string location,
            out int matchCount)
        {
            ruleId = null;
            expected = null;
            actual = null;
            location = null;
            matchCount = 0;
            try
            {
                Word.Comments comments = selection.Document.Comments;
                var matches = new List<string[]>();
                for (int i = 1; i <= comments.Count; i++)
                {
                    Word.Comment comment = comments[i];
                    // Range offsets are local to a Word story. A clicked comment
                    // must be compared with its comment text, never with body offsets.
                    Word.Range scope = selection.StoryType == Word.WdStoryType.wdCommentsStory
                        ? comment.Range : comment.Scope;
                    if (selection.StoryType != scope.StoryType)
                    {
                        continue;
                    }
                    bool intersects = selection.Start == selection.End
                        ? selection.Start >= scope.Start && selection.Start <= scope.End
                        : selection.Start < scope.End && scope.Start < selection.End;
                    if (!intersects)
                    {
                        continue;
                    }

                    // Position markers are hidden; retrieval defaults to Word's view setting.
                    // Read them explicitly without changing ShowHiddenText or the document.
                    Word.Range commentText = comment.Range.Duplicate;
                    commentText.TextRetrievalMode.IncludeHiddenText = true;
                    string text = commentText.Text ?? string.Empty;
                    Match ruleMatch = RuleIdPattern.Match(text);
                    Match expectedMatch = ExpectedValuePattern.Match(text);
                    Match actualMatch = ActualValuePattern.Match(text);
                    Match locationMatch = CommentLocationPattern.Match(text);
                    if (ruleMatch.Success && expectedMatch.Success && actualMatch.Success)
                    {
                        matches.Add(new[]
                        {
                            ruleMatch.Groups["id"].Value.Trim(),
                            expectedMatch.Groups["value"].Value.Trim(),
                            actualMatch.Groups["value"].Value.Trim(),
                            locationMatch.Success ? locationMatch.Groups["location"].Value.Trim() : string.Empty
                        });
                    }
                }

                matchCount = matches.Count;
                if (matches.Count == 0)
                {
                    return false;
                }

                // 多条批注可能共用同一个高亮范围（例如同一段同时有字体和字号问题）。
                // RangeFromPoint 无法区分重叠锚点，先激活第一条；用户可用紧凑导航切换，
                // 或展开可选的问题列表选择同一位置的其他建议。
                ruleId = matches[0][0];
                expected = matches[0][1];
                actual = matches[0][2];
                location = matches[0][3];
                return true;
            }
            catch (Exception)
            {
                // 选区或批注可能已被关闭，忽略本次联动。
                return false;
            }
        }

        /// <summary>
        /// 为触发操作的窗格所绑定的 Word 窗口生成当前内存内容的快照。
        /// 从 Word 当前内存内容生成 Flat OPC，再组装为独立 DOCX；绝不退回磁盘旧版本。
        /// </summary>
        private bool TryCreateWorkingCopy(DiagnosticsPane pane, out DocumentOperationContext operation, out string error)
        {
            operation = null;
            error = null;
            Word.Document document;
            Word.Window window;
            if (!TryGetDocumentForPane(pane, out document, out window, out error))
            {
                return false;
            }

            try
            {
                int saveFormat = (int)document.SaveFormat;
                if (saveFormat != 12 && saveFormat != 16 && saveFormat != 24)
                {
                    error = "当前仅支持 Word .docx 文档。请先另存为 .docx 后再分析或修复。";
                    return false;
                }

                string path = CreateTemporaryDocxPath("snapshot");
                try
                {
                    SaveDocumentCopy(document, path);
                    if (!File.Exists(path) || new FileInfo(path).Length == 0)
                    {
                        throw new IOException("Word 没有生成有效的当前内容副本。");
                    }

                    string fingerprint = DocxSnapshotFingerprint.Compute(path);
                    operation = new DocumentOperationContext
                    {
                        Document = document,
                        Window = window,
                        WindowHandle = window.Hwnd,
                        DocumentName = document.Name,
                        DocumentFullName = document.FullName,
                        SnapshotPath = path,
                        Fingerprint = fingerprint
                    };
                    return true;
                }
                catch
                {
                    TryDelete(path);
                    throw;
                }
            }
            catch (Exception exception)
            {
                WriteStartupLog("创建 Word 内存快照失败：" + exception);
                error = "无法可靠读取当前 Word 内容，已停止本次操作。请确认文档可另存为 .docx 后重试。详情：" + exception.Message;
                return false;
            }
        }

        private bool TryGetDocumentForPane(
            DiagnosticsPane pane,
            out Word.Document document,
            out Word.Window window,
            out string error)
        {
            document = null;
            window = null;
            error = null;
            if (pane == null || pane.IsDisposed)
            {
                error = "当前检查面板已关闭，请重新打开后重试。";
                return false;
            }

            for (int i = 0; i < CustomTaskPanes.Count; i++)
            {
                try
                {
                    CustomTaskPane taskPane = CustomTaskPanes[i];
                    if (!object.ReferenceEquals(taskPane.Control, pane))
                    {
                        continue;
                    }

                    window = taskPane.Window as Word.Window;
                    if (window == null)
                    {
                        break;
                    }

                    document = window.Document;
                    if (document != null)
                    {
                        return true;
                    }
                }
                catch (Exception exception)
                {
                    error = "无法读取此侧边栏对应的 Word 文档：" + exception.Message;
                    return false;
                }
            }

            error = "无法确认此侧边栏对应的 Word 窗口。请先激活目标文档窗口后重试。";
            return false;
        }

        private bool TryEnsureSnapshotUnchanged(DocumentOperationContext operation, out string error)
        {
            error = null;
            if (operation == null || operation.Document == null || operation.Window == null)
            {
                error = "无法确认原文窗口仍然打开，请重新分析后再操作。";
                return false;
            }

            string currentCopy = null;
            try
            {
                if (operation.Window.Hwnd != operation.WindowHandle)
                {
                    error = "原文窗口已关闭或切换，请重新分析后再操作。";
                    return false;
                }

                Word.Document windowDocument = operation.Window.Document;
                if (!string.Equals(windowDocument.Name, operation.DocumentName, StringComparison.Ordinal)
                    || !string.Equals(windowDocument.FullName, operation.DocumentFullName, StringComparison.OrdinalIgnoreCase))
                {
                    error = "原文窗口中的文档已更换，请重新分析后再操作。";
                    return false;
                }

                currentCopy = CreateTemporaryDocxPath("verify");
                SaveDocumentCopy(operation.Document, currentCopy);
                string currentFingerprint = DocxSnapshotFingerprint.Compute(currentCopy);
                if (!string.Equals(currentFingerprint, operation.Fingerprint, StringComparison.Ordinal))
                {
                    WriteStartupLog("快照差异部件：" + DocxSnapshotFingerprint.DescribeDifferences(operation.SnapshotPath, currentCopy));
                    error = "分析期间原文内容发生了变化。为避免把旧结果用于新内容，请重新分析；本次修复没有写入原文。";
                    return false;
                }

                return true;
            }
            catch (Exception exception)
            {
                error = "无法确认原文仍与分析快照一致，已停止以保护文档。请重新打开文档并分析。详情：" + exception.Message;
                return false;
            }
            finally
            {
                TryDelete(currentCopy);
            }
        }

        private static string CreateTemporaryDocxPath(string purpose)
        {
            string directory = Path.Combine(Path.GetTempPath(), "NkuThesisTool");
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, purpose + "-" + Guid.NewGuid().ToString("N") + ".docx");
        }

        private static void SaveDocumentCopy(Word.Document document, string path)
        {
            string flatOpc = document.WordOpenXML;
            FlatOpcPackageWriter.Write(flatOpc, path);
        }

        private static string CreateAnnotatedOutputPath(DocumentOperationContext operation)
        {
            string directory = null;
            string fileName = "论文";
            try
            {
                directory = operation.Document.Path;
                fileName = Path.GetFileNameWithoutExtension(operation.Document.Name);
            }
            catch (Exception)
            {
                // 未保存的新文档没有稳定目录，下面会使用用户文档目录。
            }

            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            if (string.IsNullOrWhiteSpace(fileName))
            {
                fileName = "论文";
            }

            return Path.Combine(directory, fileName + "-格式标注-"
                + DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".docx");
        }

        private Word.Document OpenAnnotatedCopy(string path, out string error)
        {
            error = null;
            try
            {
                object fileName = path;
                object confirmConversions = false;
                object readOnly = false;
                object addToRecentFiles = false;
                object passwordDocument = Type.Missing;
                object passwordTemplate = Type.Missing;
                object revert = false;
                object writePasswordDocument = Type.Missing;
                object writePasswordTemplate = Type.Missing;
                object format = Type.Missing;
                object encoding = Type.Missing;
                object visible = true;
                object openAndRepair = false;
                object documentDirection = Type.Missing;
                object noEncodingDialog = true;
                object xmlTransform = Type.Missing;
                return Application.Documents.Open(
                    ref fileName, ref confirmConversions, ref readOnly, ref addToRecentFiles,
                    ref passwordDocument, ref passwordTemplate, ref revert, ref writePasswordDocument,
                    ref writePasswordTemplate, ref format, ref encoding, ref visible, ref openAndRepair,
                    ref documentDirection, ref noEncodingDialog, ref xmlTransform);
            }
            catch (Exception exception)
            {
                error = exception.Message;
                WriteStartupLog("打开生成的 Word 副本失败：" + exception);
                return null;
            }
        }

        private static int SafeWindowHandle(Word.Window window)
        {
            try { return window != null ? window.Hwnd : -1; }
            catch (Exception) { return -1; }
        }

        private static bool PathsEqual(string first, string second)
        {
            if (string.IsNullOrWhiteSpace(first) || string.IsNullOrWhiteSpace(second))
            {
                return false;
            }

            try
            {
                return string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
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
                dialog.OverwritePrompt = false;
                while (dialog.ShowDialog() == DialogResult.OK)
                {
                    if (!File.Exists(dialog.FileName))
                    {
                        return dialog.FileName;
                    }

                    MessageBox.Show("为保护已有文件，请选择一个尚不存在的文件名。",
                        TaskPaneTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                return null;
            }
        }

        private bool EnsureGateway(DiagnosticsPane pane)
        {
            if (_gateway != null)
            {
                return true;
            }

            ShowFailure(pane, _gatewayError ?? "引擎未配置。");
            return false;
        }

        private CancellationToken GetOperationCancellationToken()
        {
            return _operationCancellation != null ? _operationCancellation.Token : CancellationToken.None;
        }

        private static bool IsPaneAvailable(DiagnosticsPane pane)
        {
            return pane != null && !pane.IsDisposed && pane.IsHandleCreated;
        }

        private static void ShowFailure(DiagnosticsPane pane, string message)
        {
            if (IsPaneAvailable(pane)) pane.ShowFailure(message);
        }

        private static void ShowAnalysis(DiagnosticsPane pane, EngineAnalysisResult result)
        {
            if (IsPaneAvailable(pane)) pane.ShowAnalysis(result);
        }

        private static void SetBusy(DiagnosticsPane pane, bool busy, string message)
        {
            if (IsPaneAvailable(pane)) pane.SetBusy(busy, message);
        }

        private static void SetStatus(DiagnosticsPane pane, string message)
        {
            if (IsPaneAvailable(pane)) pane.SetStatus(message);
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



