using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
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
    /// Word 侧入口：功能区承载操作命令，右侧问题检索列表与正文批注联动。
    ///
    /// 重要：Word 会锁住正在编辑的 .docx，引擎无法直接读取，因此所有分析/修复
    /// 都基于 Word 当前内存内容生成的临时快照。分析与修复直接更新当前打开的文档，
    /// 插件不自动保存；作者决定何时保存当前文档。
    /// </summary>
    public partial class ThisAddIn
    {
        internal const string TaskPaneTitle = "问题检索";

        /// <summary>批注正文里用于回填规则编号的片段，例如“规则：demo/east-asian-font。”。</summary>
        private static readonly Regex RuleIdPattern = new Regex(@"规则[:：]\s*(?<id>[A-Za-z0-9_\-/]+)", RegexOptions.Compiled);
        private static readonly Regex CommentLocationPattern = new Regex(
            @"\[NKU-LOCATION:(?<location>word/document\.xml#paragraph\[\d+\]/run\[\d+\])\]",
            RegexOptions.Compiled);

        private CustomTaskPane _taskPane;
        private DiagnosticsPane _pane;
        private EngineGateway _gateway;
        private ThesisRibbon _ribbon;
        private string _gatewayError;
        private System.Windows.Forms.Timer _documentHoverTimer;
        private RibbonTabVisibilityTracker _ribbonTabTracker;
        private CancellationTokenSource _operationCancellation;
        private readonly Dictionary<int, Dictionary<string, Word.Range>> _occurrenceRangesByWindow =
            new Dictionary<int, Dictionary<string, Word.Range>>();
        private int _lastHoveredWindowHandle = -1;
        private int _lastHoveredCursorX = -1;
        private int _lastHoveredCursorY = -1;
        private string _lastHoveredCommentKey;
        private bool _isNavigatingToDiagnostic;
        private bool _issuesVisible;
        private DateTime _lastNavigationUtc = DateTime.MinValue;

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
            _ribbon = new ThesisRibbon(
                OnRibbonAnalyze,
                OnRibbonFixCurrent,
                OnRibbonFixAll,
                OnRibbonVerifyResolved,
                OnRibbonNavigateIssue,
                OnRibbonNavigateOccurrence,
                SetDiagnosticsPaneVisible,
                IsDiagnosticsPaneVisible,
                () => _pane?.CanFixSelected == true,
                () => _pane?.HasFixableDiagnostics == true,
                () => _pane?.CanResolveSelected == true,
                () => _pane?.CanNavigateIssues == true,
                () => _pane?.CanNavigateOccurrences == true);
            return _ribbon;
        }

        private void SetDiagnosticsPaneVisible(bool visible)
        {
            try
            {
                Word.Window window = Application.ActiveWindow;
                EnsureTaskPaneForWindow(window);
                CustomTaskPane taskPane = FindTaskPaneForWindow(window);
                if (taskPane == null)
                {
                    throw new InvalidOperationException("当前 Word 窗口没有可用的检查面板。");
                }

                taskPane.Visible = visible;
                WriteStartupLog(visible ? "已通过功能区打开当前文档的问题检索列表。" : "已通过功能区关闭当前文档的问题检索列表。");
            }
            catch (Exception exception)
            {
                WriteStartupLog("通过功能区切换问题检索列表失败：" + exception);
                MessageBox.Show("未能切换问题检索列表。请先打开 Word 文档，再重试。",
                    TaskPaneTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private bool IsDiagnosticsPaneVisible() => _issuesVisible;

        private CustomTaskPane FindTaskPaneForWindow(Word.Window window)
        {
            if (window == null) return null;
            int windowHandle = window.Hwnd;
            for (int i = 0; i < CustomTaskPanes.Count; i++)
            {
                CustomTaskPane taskPane = CustomTaskPanes[i];
                try
                {
                    Word.Window paneWindow = taskPane.Window as Word.Window;
                    if (paneWindow != null && paneWindow.Hwnd == windowHandle) return taskPane;
                }
                catch (Exception)
                {
                    // Skip panes whose Word window is closing.
                }
            }
            return null;
        }

        private void OnRibbonAnalyze()
        {
            DiagnosticsPane pane = GetPaneForRibbonCommand();
            if (pane != null) OnAnalyzeRequested(pane, EventArgs.Empty);
        }

        private void OnRibbonFixCurrent()
        {
            DiagnosticsPane pane = GetPaneForRibbonCommand();
            if (pane != null) OnFixCurrentRequested(pane, EventArgs.Empty);
        }

        private void OnRibbonFixAll()
        {
            DiagnosticsPane pane = GetPaneForRibbonCommand();
            if (pane != null) OnFixAllRequested(pane, EventArgs.Empty);
        }

        private void OnRibbonVerifyResolved()
        {
            DiagnosticsPane pane = GetPaneForRibbonCommand();
            if (pane != null) OnVerifyResolvedRequested(pane, EventArgs.Empty);
        }

        private void OnRibbonNavigateIssue(int direction)
        {
            DiagnosticsPane pane = GetPaneForRibbonCommand();
            pane?.NavigateIssue(direction);
        }

        private void OnRibbonNavigateOccurrence(int direction)
        {
            DiagnosticsPane pane = GetPaneForRibbonCommand();
            pane?.NavigateOccurrence(direction);
        }

        private DiagnosticsPane GetPaneForRibbonCommand()
        {
            try
            {
                return GetPaneForWindow(Application.ActiveWindow);
            }
            catch (Exception exception)
            {
                WriteStartupLog("功能区命令未执行：当前没有可用的 Word 文档窗口。" + exception.Message);
                return null;
            }
        }

        private DiagnosticsPane TryGetPaneForActiveWindow()
        {
            try
            {
                Word.Window activeWindow = Application.ActiveWindow;
                if (activeWindow == null) return null;
                int windowHandle = activeWindow.Hwnd;
                for (int i = 0; i < CustomTaskPanes.Count; i++)
                {
                    CustomTaskPane taskPane = CustomTaskPanes[i];
                    Word.Window paneWindow = taskPane.Window as Word.Window;
                    if (paneWindow != null && paneWindow.Hwnd == windowHandle)
                        return taskPane.Control as DiagnosticsPane;
                }
            }
            catch (Exception)
            {
                // Word's backstage page and closing-window transitions have no active pane.
            }
            return null;
        }

        private void OnPaneStateChanged(object sender, EventArgs e)
        {
            _ribbon?.Invalidate();
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
            _ribbonTabTracker = new RibbonTabVisibilityTracker();

            // Word 没有文档正文的鼠标悬停事件。用 UI 计时器查询鼠标下的
            // Word.Range，再通过批注锚点选择问题列表项；COM 调用仍在 Word UI 线程执行。
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
                + "；问题检索列表：" + (_taskPane != null ? "已创建" : "等待文档窗口"));
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
            _ribbonTabTracker?.Dispose();
            _ribbonTabTracker = null;
            _occurrenceRangesByWindow.Clear();
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
            string snapshotPath = null;
            try
            {
                WriteStartupLog("收到“分析当前文档”点击。");
                if (!EnsureGateway(pane))
                {
                    return;
                }
                EngineGateway gateway = _gateway;
                CancellationToken cancellationToken = GetOperationCancellationToken();

                DocumentOperationContext operation;
                string failure;
                if (!TryCreateSnapshot(pane, out operation, out failure))
                {
                    ShowFailure(pane, failure);
                    return;
                }
                snapshotPath = operation.SnapshotPath;

                WriteStartupLog("已为 Word 窗口 " + operation.WindowHandle + " 生成当前内容快照，开始分析。");
                SetBusy(pane, true, "正在分析当前文档并更新本页标注…");
                EngineResponse response = await gateway.AnalyzeAsync(snapshotPath, cancellationToken);
                if (response == null || !response.Ok || response.Result == null)
                {
                    string analysisError = EngineGateway.DescribeError(response != null ? response.Error : null);
                    WriteStartupLog("引擎分析返回错误：" + analysisError);
                    ShowFailure(pane, analysisError);
                    return;
                }

                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    ShowFailure(pane, failure);
                    return;
                }

                string syncError = null;
                bool synchronized = pane.InvokeOnUiThread(() => WordDocumentSynchronizer.RefreshAnnotations(
                    operation.Document, operation.SnapshotPath, response.Result, out syncError));
                if (!synchronized)
                {
                    WriteStartupLog("刷新当前文档格式标记失败：" + syncError);
                    ShowFailure(pane, "无法更新当前文档的格式标注：" + syncError);
                    return;
                }

                ShowAnalysis(pane, response.Result);
                pane.InvokeOnUiThread(() => CacheOccurrenceRanges(operation.WindowHandle, operation.Document,
                    operation.SnapshotPath, response.Result));
                SetStatus(pane, "分析完成：当前文档共 " + response.Result.Diagnostics.Count
                    + " 项问题，原文黄色标记已在本窗口更新。可用 Ctrl+Z 撤销标记；完成后自行保存。");
                WriteStartupLog("已在当前 Word 文档刷新标注，问题数 " + response.Result.Diagnostics.Count
                    + "，窗口 " + operation.WindowHandle + "；未打开或保存其他文档。");
            }
            catch (Exception exception)
            {
                WriteStartupLog("分析出现异常：" + exception);
                ShowFailure(pane, "分析出现异常，已停止：" + exception.Message);
            }
            finally
            {
                SetBusy(pane, false, null);
                TryDelete(snapshotPath);
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

            string status = NavigateTo(_pane, diagnostic, _pane.SelectedLocation);
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
                SetStatus(pane, "请先从右侧问题列表或 Word 批注选择一个问题，再执行“修复当前问题”。");
                return;
            }

            if (!string.Equals(diagnostic.FixLevel, "auto", StringComparison.OrdinalIgnoreCase))
            {
                SetStatus(pane, "该问题不能自动修复（处理方式：" + DiagnosticsPane.DescribeFixLevel(diagnostic.FixLevel)
                    + "）。请按说明手动调整，或先与规则负责人确认后再纳入自动修复。");
                return;
            }

            await RunRepair(pane, diagnostic, "当前选中项 " + DiagnosticsPane.DescribeRule(diagnostic.RuleId));
        }

        private async void OnFixAllRequested(object sender, EventArgs e)
        {
            ActivatePaneFromSender(sender);
            WriteStartupLog("收到“预览并修复全部”点击。");
            await RunRepair(_pane, null, "全部可自动修复项");
        }

        private async void OnVerifyResolvedRequested(object sender, EventArgs e)
        {
            ActivatePaneFromSender(sender);
            DiagnosticsPane pane = sender as DiagnosticsPane ?? _pane;
            EngineDiagnostic selected = pane?.SelectedDiagnostic;
            string location = pane?.SelectedLocation;
            if (selected == null || string.IsNullOrWhiteSpace(location))
            {
                SetStatus(pane, "请先从问题列表选择一处，再验证是否已解决。");
                return;
            }

            string snapshotPath = null;
            try
            {
                if (!EnsureGateway(pane)) return;
                DocumentOperationContext operation;
                string failure;
                if (!TryCreateSnapshot(pane, out operation, out failure))
                {
                    ShowFailure(pane, failure);
                    return;
                }
                snapshotPath = operation.SnapshotPath;
                SetBusy(pane, true, "正在重新检查所选位置…");
                EngineResponse response = await _gateway.AnalyzeAsync(snapshotPath, GetOperationCancellationToken());
                if (response == null || !response.Ok || response.Result == null)
                {
                    ShowFailure(pane, EngineGateway.DescribeError(response?.Error));
                    return;
                }
                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    ShowFailure(pane, failure);
                    return;
                }

                string syncError = null;
                bool synchronized = pane.InvokeOnUiThread(() => WordDocumentSynchronizer.RefreshAnnotations(
                    operation.Document, operation.SnapshotPath, response.Result, out syncError));
                if (!synchronized)
                {
                    ShowFailure(pane, "无法刷新当前文档标记：" + syncError);
                    return;
                }

                bool remains = response.Result.Diagnostics.Any(item =>
                    string.Equals(item.RuleId, selected.RuleId, StringComparison.Ordinal) &&
                    string.Equals(item.Expected, selected.Expected, StringComparison.Ordinal) &&
                    HasOccurrence(item, location));
                ShowAnalysis(pane, response.Result);
                pane.InvokeOnUiThread(() => CacheOccurrenceRanges(operation.WindowHandle, operation.Document,
                    operation.SnapshotPath, response.Result));
                if (remains)
                {
                    pane.SelectByOccurrenceLocation(location, navigateToSource: false);
                    SetStatus(pane, "重新检查后，该位置仍不符合要求，问题和黄色标记已保留。修改后可再次验证。");
                }
                else
                {
                    SetStatus(pane, "已重新检查：所选位置不再触发该项规则，黄色标记已清除。改动仍在当前文档中，请由你决定是否保存。");
                }
            }
            catch (Exception exception)
            {
                ShowFailure(pane, "验证解决状态失败：" + exception.Message);
            }
            finally
            {
                SetBusy(pane, false, null);
                TryDelete(snapshotPath);
            }
        }

        private static bool HasOccurrence(EngineDiagnostic diagnostic, string location)
        {
            if (diagnostic == null) return false;
            if (diagnostic.Occurrences != null && diagnostic.Occurrences.Count > 0)
                return diagnostic.Occurrences.Any(item => string.Equals(item?.Location, location, StringComparison.Ordinal));
            return string.Equals(diagnostic.Location, location, StringComparison.Ordinal);
        }

        private static string DescribeLocation(string location)
        {
            DocumentLocation parsed;
            string error;
            return DocumentLocator.TryParse(location, out parsed, out error)
                ? DocumentLocator.Describe(parsed) : location;
        }

        /// <summary>
        /// Word 每个文档窗口都需要自己的任务窗格和诊断列表。启动时只创建首个窗口的
        /// 窗格会导致后来打开的文档没有侧栏；WindowActivate 同时覆盖新窗口和窗口切换。
        /// </summary>
        private void OnWindowActivate(Word.Document document, Word.Window window)
        {
            EnsureTaskPaneForWindow(window);
            _issuesVisible = _taskPane != null && _taskPane.Visible;
            _ribbon?.Invalidate();
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
            int preferredWidth = pane.ConfigureForOfficeWindow(windowHandle);
            CustomTaskPane taskPane = null;
            try
            {
                taskPane = CustomTaskPanes.Add(pane, TaskPaneTitle, window);
                // The right pane is only a searchable issue index; all actions live on the ribbon.
                taskPane.Width = preferredWidth;
                taskPane.VisibleChanged += OnTaskPaneVisibleChanged;
                taskPane.Visible = false;
                _taskPane = taskPane;
                _pane = pane;
                pane.CreateControl();
                CustomTaskPane trackedTaskPane = taskPane;
                _ribbonTabTracker?.Track(windowHandle, selected =>
                {
                    if (pane.IsDisposed || !pane.IsHandleCreated) return;
                    try
                    {
                        pane.BeginInvoke(new Action(() =>
                        {
                            if (pane.IsDisposed) return;
                            try
                            {
                                if (trackedTaskPane.Visible != selected) trackedTaskPane.Visible = selected;
                            }
                            catch (Exception) { /* The document window may be closing. */ }
                        }));
                    }
                    catch (InvalidOperationException) { /* The pane has closed. */ }
                });
                pane.SetStatus(string.IsNullOrEmpty(_gatewayError)
                    ? "点击功能区“分析当前文档”开始检查。"
                    : _gatewayError);
                WriteStartupLog("已为 Word 窗口 " + windowHandle + " 创建问题检索列表。");
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
                        // Keep the user's explicit open/closed choice. Ribbon commands can
                        // still work against the pane's state while the search list is hidden.
                        return taskPane.Control as DiagnosticsPane;
                    }
                }
            }
            catch (Exception exception)
            {
                WriteStartupLog("无法取得当前文档的侧边栏：" + exception.Message);
            }

            return null;
        }

        private DiagnosticsPane CreateDiagnosticsPane()
        {
            var pane = new DiagnosticsPane();
            pane.DiagnosticSelected += OnDiagnosticSelected;
            pane.StateChanged += OnPaneStateChanged;
            return pane;
        }

        private void OnTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            if (sender is CustomTaskPane taskPane && object.ReferenceEquals(taskPane, _taskPane))
                _issuesVisible = taskPane.Visible;
            _ribbon?.Invalidate();
        }

        private void ActivatePaneFromSender(object sender)
        {
            DiagnosticsPane pane = sender as DiagnosticsPane;
            if (pane == null)
            {
                return;
            }

            if (object.ReferenceEquals(_pane, pane) && _taskPane != null) return;

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
        /// 修复闭环：临时快照 → 生成计划 → 中文预览确认 → 修改当前 Word 文档 → 重新分析。
        /// 修改留在当前打开的文档中，由作者决定何时保存。
        /// </summary>
        private async Task RunRepair(DiagnosticsPane pane, EngineDiagnostic diagnostic, string scopeLabel)
        {
            string workCopy = null;
            string applyOutput = null;
            string finalOutput = null;
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
                if (!TryCreateSnapshot(pane, out operation, out failure))
                {
                    WriteStartupLog("修复中止：临时副本失败 - " + failure);
                    ShowFailure(pane, failure);
                    return;
                }
                workCopy = operation.SnapshotPath;

                if (!pane.InvokeOnUiThread(() => HasAnalysisMarker(operation.Document)))
                {
                    ShowFailure(pane, "请先分析当前文档，再在本窗口内修复。");
                    return;
                }

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
                        return preview.ShowForOfficeWindow(operation.WindowHandle);
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

                applyOutput = CreateTemporaryDocxPath("repair");

                /*
                 * The dialogs above are synchronous UI-thread calls. After they return,
                 * the engine can continue on a worker thread; DiagnosticsPane marshals
                 * subsequent status and result updates back to the same UI thread.
                 */
                WriteStartupLog("开始生成修复中间副本。");
                SetStatus(pane, "正在更新当前文档…");
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

                finalOutput = CreateTemporaryDocxPath("refreshed");
                SetStatus(pane, "正在复查修复结果并刷新当前文档标记；长文档可能需要一两分钟…");
                WriteStartupLog("开始重新分析修复结果并刷新标记；剩余待处理 "
                    + (plan.Deferred != null ? plan.Deferred.Count : 0) + " 项。");
                EngineResponse refreshed = await gateway.AnnotateAsync(applied.OutputPath, finalOutput, cancellationToken);
                if (refreshed == null || !refreshed.Ok || refreshed.Result == null || refreshed.Annotation == null
                    || !PathsEqual(refreshed.Annotation.OutputPath, finalOutput)
                    || !File.Exists(finalOutput))
                {
                    string refreshError = EngineGateway.DescribeError(refreshed != null ? refreshed.Error : null);
                    if (refreshed != null && refreshed.Ok && refreshed.Annotation != null
                        && !PathsEqual(refreshed.Annotation.OutputPath, finalOutput))
                    {
                        refreshError = "标注引擎返回了与请求不一致的输出路径。";
                    }
                    WriteStartupLog("修复结果重新分析或标注失败；Word 原文尚未写入：" + refreshError);
                    ShowFailure(pane, applySummary + Environment.NewLine + "当前文档尚未修改；刷新批注失败：" + refreshError);
                    return;
                }

                WriteStartupLog("修复结果重新分析完成：剩余问题 " + refreshed.Result.Diagnostics.Count
                    + " 项；将刷新 " + refreshed.Annotation.CommentCount + " 条汇总批注。");

                if (!TryEnsureSnapshotUnchanged(operation, out failure))
                {
                    TryDelete(finalOutput);
                    finalOutput = null;
                    ShowFailure(pane, failure);
                    return;
                }

                string syncError = null;
                bool synchronized = pane.InvokeOnUiThread(() => WordDocumentSynchronizer.ApplyRepairsAndRefresh(
                    operation.Document, operation.SnapshotPath, applied.Applied, refreshed.Result, out syncError));
                if (!synchronized)
                {
                    WriteStartupLog("Word 原文修复或标注同步失败：" + syncError);
                    ShowFailure(pane, applySummary + Environment.NewLine + "Word 当前文档未完成修复："
                        + syncError + Environment.NewLine + "请检查文档状态后重新分析。当前文档：" + operation.DocumentFullName);
                    return;
                }

                Word.Document resultDocument = operation.Document;
                DiagnosticsPane resultPane = GetPaneForWindow(resultDocument.ActiveWindow);
                ShowAnalysis(resultPane, refreshed.Result);
                pane.InvokeOnUiThread(() => CacheOccurrenceRanges(operation.WindowHandle, resultDocument,
                    operation.SnapshotPath, refreshed.Result));
                SetStatus(resultPane, applySummary + Environment.NewLine
                    + "修复已应用到当前 Word 文档，批注已刷新；剩余 " + refreshed.Result.Diagnostics.Count + " 项。\n"
                    + "改动尚未保存，可用 Ctrl+Z 撤销；完成后按 Ctrl+S 保存当前文档。\n"
                    + operation.DocumentFullName);
                SetStatus(pane, "修复完成：已直接修改当前 Word 文档，没有另建结果文档；需要时按 Ctrl+S 保存。\n"
                    + operation.DocumentFullName);
                WriteStartupLog("已在当前 Word 文档中直接修复并刷新批注；剩余问题 "
                    + refreshed.Result.Diagnostics.Count + " 项，等待作者保存。");
            }
            catch (Exception exception)
            {
                WriteStartupLog("修复过程出现异常：" + exception);
                ShowFailure(pane, "修复过程出现异常，已停止：" + exception.Message);
            }
            finally
            {
                SetBusy(pane, false, null);
                WriteStartupLog("修复操作结束；操作按钮已恢复。");
                TryDelete(workCopy);
                TryDelete(applyOutput);
                TryDelete(finalOutput);
            }
        }

        private string ReportApplyResult(EngineApplyResult result)
        {
            var message = new System.Text.StringBuilder();
            message.Append(result.Failures != null && result.Failures.Count > 0
                ? "引擎修复未完全应用：已准备 "
                : "引擎已准备写回 ").Append(result.Applied.Count).Append(" 处");
            if (result.Applied.Count > 0)
            {
                message.Append("（");
                string[] rules = result.Applied.Select(patch => patch.RuleId).Distinct(StringComparer.Ordinal).ToArray();
                for (int i = 0; i < rules.Length; i++)
                {
                    if (i > 0) message.Append("、");
                    message.Append(DiagnosticsPane.DescribeRule(rules[i]));
                }
                message.Append("）");
            }

            if (result.Failures.Count > 0)
            {
                message.Append("；失败 ").Append(result.Failures.Count).Append(" 处（未产生任何部分写入）");
            }

            if (string.IsNullOrWhiteSpace(result.OutputPath))
            {
                message.Append("。没有生成修复文件。");
            }
            else
            {
                message.Append("。等待同步到当前 Word 文档。");
            }
            foreach (EnginePatchFailure failure in result.Failures)
            {
                message.Append(Environment.NewLine).Append("· 失败：").Append(DiagnosticsPane.DescribeRule(failure.Patch.RuleId))
                    .Append(" @ ").Append(DescribeLocation(failure.Patch.Location)).Append(" — ").Append(failure.Reason);
            }

            return message.ToString();
        }

        /// <summary>把一条诊断定位到原文：选中对应段落（选中即为可见高亮，不修改文档内容）。</summary>
        private string NavigateTo(DiagnosticsPane pane, EngineDiagnostic diagnostic, string selectedLocation)
        {
            var navigationWatch = Stopwatch.StartNew();
            _lastNavigationUtc = DateTime.UtcNow;
            DocumentLocation location;
            string parseError;
            if (!DocumentLocator.TryParse(selectedLocation ?? diagnostic.Location, out location, out parseError))
            {
                return parseError;
            }

            Word.Document document;
            string contextError;
            Word.Window window;
            if (!TryGetDocumentForPane(pane, out document, out window, out contextError))
            {
                return contextError;
            }

            string locationKey = selectedLocation ?? diagnostic.Location;
            try
            {
                if (window == null || !_occurrenceRangesByWindow.TryGetValue(window.Hwnd,
                    out Dictionary<string, Word.Range> cachedRanges)
                    || !cachedRanges.TryGetValue(locationKey, out Word.Range cachedRange)
                    || cachedRange == null)
                    return "原文定位缓存已失效，请重新分析当前文档后再定位。";

                // 程序化选中段落会触发 WindowSelectionChange。重叠批注无法由段落选区
                // 唯一辨认，忽略这一次回调，避免导航到后续问题时又切回第一条批注。
                _isNavigatingToDiagnostic = true;
                cachedRange.Duplicate.Select();
                return "已定位到" + DocumentLocator.Describe(location) + "：" + DiagnosticsPane.DescribeRule(diagnostic.RuleId) + "。";
            }
            catch (Exception exception)
            {
                return "定位失败：" + exception.Message;
            }
            finally
            {
                _isNavigatingToDiagnostic = false;
                _lastNavigationUtc = DateTime.UtcNow;
                WriteStartupLog("问题定位耗时：" + navigationWatch.ElapsedMilliseconds + "ms；位置=" + locationKey);
            }
        }

        /// <summary>
        /// 批注锚点到问题检索项的选择联动。
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
                    ? "此处有多条批注，当前关联第一项；可用功能区上一项/下一项浏览。"
                    : "当前批注已关联到问题：" + DiagnosticsPane.DescribeRule(ruleId) + "。");
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
                // Reject task-pane motion and stationary pointers before any Word COM.
                // Explicit navigation owns the selection briefly, so hover cannot undo it.
                System.Drawing.Point cursor = Cursor.Position;
                if (_isNavigatingToDiagnostic || _pane == null || _pane.IsBusy ||
                    Control.MouseButtons != MouseButtons.None ||
                    (DateTime.UtcNow - _lastNavigationUtc).TotalMilliseconds < 700) return;
                if (_pane.IsHandleCreated && _pane.RectangleToScreen(_pane.ClientRectangle).Contains(cursor)) return;
                if (cursor.X == _lastHoveredCursorX && cursor.Y == _lastHoveredCursorY) return;

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
                if (windowHandle == _lastHoveredWindowHandle
                    && cursor.X == _lastHoveredCursorX
                    && cursor.Y == _lastHoveredCursorY)
                {
                    return;
                }

                _lastHoveredWindowHandle = windowHandle;
                _lastHoveredCursorX = cursor.X;
                _lastHoveredCursorY = cursor.Y;

                if (_pane.IsHandleCreated && _pane.RectangleToScreen(_pane.ClientRectangle).Contains(cursor))
                {
                    // The polling bridge is only for document text. Do no Word COM work
                    // while the pointer is moving through the searchable issue list.
                    return;
                }

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
                            _pane.SetStatus("此处有多条批注，已关联第一项；可用功能区上一项/下一项浏览。");
                        }
                    }
                }
                else if (TryReadHoveredOccurrence(windowHandle, range, out location))
                {
                    string occurrenceKey = "occurrence\u001f" + location;
                    if (!string.Equals(occurrenceKey, _lastHoveredCommentKey, StringComparison.Ordinal)
                        && _pane.SelectByOccurrenceLocation(location, navigateToSource: false))
                    {
                        _lastHoveredCommentKey = occurrenceKey;
                        WriteStartupLog("合并问题悬停联动：window=" + windowHandle + "; location=" + location);
                        _pane.SetStatus("当前合并问题位置：" + DescribeLocation(location) + "。可用功能区上一处/下一处浏览同一要求的其他位置。");
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

        private void CacheOccurrenceRanges(int windowHandle, Word.Document document, string snapshotPath,
            EngineAnalysisResult analysis)
        {
            try
            {
                Dictionary<string, Word.Range> ranges = WordDocumentSynchronizer.GetOccurrenceRanges(
                    document, snapshotPath, analysis);
                if (ranges.Count == 0)
                    _occurrenceRangesByWindow.Remove(windowHandle);
                else
                    _occurrenceRangesByWindow[windowHandle] = ranges;
            }
            catch (Exception exception)
            {
                _occurrenceRangesByWindow.Remove(windowHandle);
                WriteStartupLog("未能缓存问题位置，请重新分析当前文档后再定位：" + exception.Message);
            }
        }

        private bool TryReadHoveredOccurrence(int windowHandle, Word.Range hoveredRange, out string location)
        {
            location = null;
            if (hoveredRange == null || !_occurrenceRangesByWindow.TryGetValue(windowHandle, out Dictionary<string, Word.Range> ranges))
                return false;

            try
            {
                int start = hoveredRange.Start;
                int end = hoveredRange.End;
                foreach (KeyValuePair<string, Word.Range> target in ranges)
                {
                    if (target.Value == null || target.Value.StoryType != hoveredRange.StoryType)
                        continue;
                    bool intersects = start == end
                        ? start >= target.Value.Start && start < target.Value.End
                        : start < target.Value.End && target.Value.Start < end;
                    if (intersects)
                    {
                        location = target.Key;
                        return true;
                    }
                }
            }
            catch (Exception)
            {
                _occurrenceRangesByWindow.Remove(windowHandle);
            }
            return false;
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
                    string expectedValue = CommentDiagnosticMatcher.ReadCommentField(text, "期望值");
                    string actualValue = CommentDiagnosticMatcher.ReadCommentField(text, "实际值");
                    Match locationMatch = CommentLocationPattern.Match(text);
                    if (ruleMatch.Success && expectedValue.Length > 0 && actualValue.Length > 0)
                    {
                        matches.Add(new[]
                        {
                            ruleMatch.Groups["id"].Value.Trim(),
                            expectedValue,
                            actualValue,
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
        private bool TryCreateSnapshot(DiagnosticsPane pane, out DocumentOperationContext operation, out string error)
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
                if (document.ReadOnly || document.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    error = "当前文档为只读或受保护，无法在本页写入标注或修复。请先启用编辑后重试。";
                    return false;
                }

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

        private static bool HasAnalysisMarker(Word.Document document)
        {
            if (document == null)
            {
                return false;
            }

            try
            {
                if (WordDocumentSynchronizer.IsWorkingCopy(document))
                    return true;
                for (int i = 1; i <= document.Comments.Count; i++)
                {
                    Word.Comment comment = document.Comments[i];
                    Word.Range text = comment.Range.Duplicate;
                    text.TextRetrievalMode.IncludeHiddenText = true;
                    if ((text.Text ?? string.Empty).IndexOf("[NKU-TOOL-ANNOTATION:v1]", StringComparison.Ordinal) >= 0)
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
                // If marker inspection fails, require analysis again before repair.
                // Do not infer permission to repair from an unreadable marker.
            }

            return false;
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



