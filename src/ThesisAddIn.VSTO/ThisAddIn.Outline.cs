using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ThesisAddIn.Protocol;
using ThesisAddIn.VSTO.Engine;
using ThesisAddIn.VSTO.Ui;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO
{
    public partial class ThisAddIn
    {
        private sealed class OutlineHistory
        {
            internal Word.Document Document;
            internal WordDocumentSynchronizer.OutlineFormattingRecord Record;
        }
        private readonly Dictionary<int, OutlineHistory> _outlineHistory = new Dictionary<int, OutlineHistory>();
        private bool _outlineCommandQueued;
        private sealed class OutlineUiContext : SynchronizationContext
        {
            private readonly Control _owner;
            internal OutlineUiContext(Control owner) { _owner = owner; }
            public override SynchronizationContext CreateCopy() => this;
            public override void Post(SendOrPostCallback callback, object state) =>
                _owner.BeginInvoke(new Action(() => Dispatch(callback, state)));
            public override void Send(SendOrPostCallback callback, object state)
            {
                if (_owner.InvokeRequired) _owner.Invoke(new Action(() => Dispatch(callback, state)));
                else Dispatch(callback, state);
            }
            private void Dispatch(SendOrPostCallback callback, object state)
            {
                var previous = Current;
                SetSynchronizationContext(this);
                try { callback(state); } finally { SetSynchronizationContext(previous); }
            }
        }
        private void QueueOutlineCommand(Func<DiagnosticsPane, Task> command)
        {
            DiagnosticsPane pane = GetPaneForRibbonCommand();
            if (pane == null || pane.IsBusy || _outlineCommandQueued || !EnsureGateway(pane)) return;
            _outlineCommandQueued = true;
            SetBusy(pane, true, "正在准备整篇排版操作…");
            try {
                pane.BeginInvoke(new Action(async () => {
                    var previous = SynchronizationContext.Current;
                    SynchronizationContext.SetSynchronizationContext(new OutlineUiContext(pane));
                    try {
                        AssertOutlineUiThread(pane);
                        WriteStartupLog("整篇操作 UI thread=" + Thread.CurrentThread.ManagedThreadId + "; apartment=" +
                            Thread.CurrentThread.GetApartmentState() + "; context=" + SynchronizationContext.Current.GetType().Name);
                        await command(pane);
                    } catch (Exception exception) { WriteStartupLog("整篇操作调度失败：" + exception); ShowFailure(pane, exception.Message); }
                    finally { SynchronizationContext.SetSynchronizationContext(previous); _outlineCommandQueued = false;
                        SetBusy(pane, false, null); _ribbon?.Invalidate(); }
                }));
            } catch { _outlineCommandQueued = false; SetBusy(pane, false, null); throw; }
        }
        private static void AssertOutlineUiThread(DiagnosticsPane pane)
        {
            if (pane.InvokeRequired || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("整篇排版必须在 Word 的 UI STA 线程运行，已停止操作。");
        }
        // Embedded Office interfaces can produce different RCW references for
        // the same open document after a modal/await. Word keeps Name+FullName
        // unique among its open documents; closed wrappers throw and are rejected.
        private static bool SameOutlineDocument(Word.Document first, Word.Document second)
        {
            if (first == null || second == null) return false;
            try { return string.Equals(first.Name, second.Name, StringComparison.Ordinal) &&
                string.Equals(first.FullName, second.FullName, StringComparison.OrdinalIgnoreCase); }
            catch { return false; }
        }
        private bool CanReviewOutline() => _pane != null && !_pane.IsBusy && !_outlineCommandQueued && _gateway != null;
        private bool CanUndoOutline()
        {
            try {
                if (_pane == null || _pane.IsBusy || _outlineCommandQueued) return false;
                Word.Window window = Application.ActiveWindow;
                return _outlineHistory.TryGetValue(window.Hwnd, out var history) && SameOutlineDocument(history.Document, window.Document) && history.Record.Count > 0;
            } catch { return false; }
        }

        private void OnRibbonReviewOutline() => QueueOutlineCommand(ReviewOutlineCoreAsync);
        private async Task ReviewOutlineCoreAsync(DiagnosticsPane pane)
        {
            DocumentOperationContext operation = null;
            WordDocumentSynchronizer.OutlineFormattingRecord applied = null;
            bool verified = false;
            try
            {
                ClearHoverPreview();
                SetBusy(pane, true, "正在识别整篇结构与安全原文位置…");
                string error;
                if (!TryCreateSnapshot(pane, out operation, out error)) throw new InvalidDataException(error);
                if (operation.Document.TrackRevisions || operation.Document.Revisions.Count != 0)
                    throw new InvalidDataException("请先由作者处理已有修订并关闭修订模式，再进行整篇排版。");
                var response = await _gateway.ReviewStructureAsync(operation.SnapshotPath, GetOperationCancellationToken());
                AssertOutlineUiThread(pane);
                if (response == null || !response.Ok || response.StructureReview == null)
                    throw new InvalidDataException(EngineGateway.DescribeError(response?.Error));
                if (!TryEnsureSnapshotUnchanged(operation, out error)) throw new InvalidDataException(error);
                EngineStructurePlanRequest request;
                using (var dialog = new ThesisOutlineReviewDialog(response.StructureReview))
                {
                    dialog.LocateRequested += block => {
                        try { WordDocumentSynchronizer.LocateOutlineBlock(operation.Document, operation.SnapshotPath, block); }
                        catch (Exception exception) { SetStatus(pane, "大纲定位未完成：" + exception.Message); }
                    };
                    if (dialog.ShowForOfficeWindow(operation.WindowHandle) != DialogResult.OK) return;
                    request = new EngineStructurePlanRequest {
                        DocumentFingerprint = response.StructureReview.DocumentFingerprint,
                        PackageSha256 = response.StructureReview.PackageSha256,
                        ApprovedBlocks = dialog.ConfirmedRoles.ToList(), TitleMode = dialog.TitleMode,
                        AuthorSelectedDraft = true
                    };
                }
                if (!TryEnsureSnapshotUnchanged(operation, out error)) throw new InvalidDataException(error);
                SetStatus(pane, "正在匹配确认的大纲与南开 2026 格式配置…");
                var planned = await _gateway.PlanStructureAsync(operation.SnapshotPath, request, GetOperationCancellationToken());
                if (planned == null || !planned.Ok || planned.StructurePlan == null)
                    throw new InvalidDataException(EngineGateway.DescribeError(planned?.Error));
                EngineStructureFormatPlan plan = planned.StructurePlan;
                if (plan.Operations.Count == 0) throw new InvalidDataException("所选角色暂未覆盖可自动排版的格式；原文未修改。");
                if (!TryEnsureSnapshotUnchanged(operation, out error)) throw new InvalidDataException(error);
                var watch = Stopwatch.StartNew();
                applied = new WordDocumentSynchronizer.OutlineFormattingRecord();
                await WordDocumentSynchronizer.ApplyOutlineFormattingAsync(operation.Document,
                    operation.SnapshotPath, plan, () => {
                        AssertOutlineUiThread(pane);
                        GetOperationCancellationToken().ThrowIfCancellationRequested();
                        if (operation.Window.Hwnd != operation.WindowHandle || !SameOutlineDocument(operation.Window.Document, operation.Document) ||
                            !SameOutlineDocument(Application.ActiveDocument, operation.Document))
                            throw new InvalidDataException("活动文档窗口已改变，已停止整篇排版。");
                    }, (count, total) => SetStatus(pane, count == 0 ? "正在核对段落和字符位置…" :
                        "正在排版已确认段落 " + count + "/" + total + "…"), applied, () => {
                            if (!TryEnsureSnapshotUnchanged(operation, out string changed)) throw new InvalidDataException(changed);
                        });

                // The host reopens a fresh snapshot and checks native text anchors,
                // canonical targets and effective formatting after Word serialization.
                string checkPath = CreateTemporaryDocxPath("outline-format-readback");
                try {
                    SaveDocumentCopy(operation.Document, checkPath);
                    string checkFingerprint = DocxSnapshotFingerprint.Compute(checkPath);
                    var checkedResponse = await _gateway.VerifyStructureFormatAsync(checkPath, plan, GetOperationCancellationToken());
                    if (checkedResponse == null || !checkedResponse.Ok || checkedResponse.StructureVerification == null ||
                        !checkedResponse.StructureVerification.Succeeded)
                    {
                        string details = checkedResponse?.StructureVerification == null
                            ? EngineGateway.DescribeError(checkedResponse?.Error)
                            : string.Join("；", checkedResponse.StructureVerification.Issues.Take(3).Select(item => item.Reason + "（" + item.Property + "）"));
                        throw new InvalidDataException("整篇排版回读检查未通过：" + details);
                    }
                    string latest = CreateTemporaryDocxPath("outline-readback-current");
                    try {
                        SaveDocumentCopy(operation.Document, latest);
                        if (DocxSnapshotFingerprint.Compute(latest) != checkFingerprint)
                            throw new InvalidDataException("回读检查期间文档被再次编辑，已停止以保护后续改动。");
                    } finally { TryDelete(latest); }
                } finally { TryDelete(checkPath); }
                verified = true;
                _outlineHistory[operation.WindowHandle] = new OutlineHistory { Document = operation.Document, Record = applied };
                GetHistory(pane).Clear();
                if (_repairBasis.TryGetValue(operation.WindowHandle, out var oldBasis)) TryDelete(oldBasis.MappingPath);
                _repairBasis.Remove(operation.WindowHandle);
                await RefreshReviewFromLive(pane, operation.Document, operation.Window);
                string completion = "整篇排版完成：" + applied.Count + " 段通过格式回读；" + plan.Omissions.Count +
                    " 个块保留原格式（未勾选或未覆盖）。未自动保存，可用“撤销整篇排版”恢复。";
                SetStatus(pane, completion);
                WriteStartupLog(completion + "；写入、回读及重分析耗时 " + watch.ElapsedMilliseconds + "ms。");
            }
            catch (Exception exception)
            {
                string message = exception.Message;
                if (!verified && applied != null && applied.Count > 0 && operation != null)
                {
                    var rollback = await WordDocumentSynchronizer.UndoOutlineFormattingAsync(operation.Document, applied, () => { });
                    if (!rollback.Item1) {
                        message += "；不能自动恢复：" + rollback.Item2;
                        _outlineHistory[operation.WindowHandle] = new OutlineHistory { Document = operation.Document, Record = applied };
                    }
                    else message += "；本轮已写入格式已恢复。";
                }
                string prefix = verified ? "整篇排版已通过回读，刷新问题列表未完成：" : "整篇排版未完成：";
                WriteStartupLog(prefix + exception);
                ShowFailure(pane, prefix + message);
                MessageBox.Show(prefix + message, "大纲确认与排版", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally { SetBusy(pane, false, null); TryDelete(operation?.SnapshotPath); _ribbon?.Invalidate(); }
        }

        private void OnRibbonUndoOutline() => QueueOutlineCommand(UndoOutlineCoreAsync);
        private async Task UndoOutlineCoreAsync(DiagnosticsPane pane)
        {
            try {
                ClearHoverPreview();
                Word.Window window = Application.ActiveWindow;
                if (!_outlineHistory.TryGetValue(window.Hwnd, out var history) || !SameOutlineDocument(history.Document, window.Document))
                    throw new InvalidDataException("当前文档没有整篇排版撤销记录。");
                SetBusy(pane, true, "正在核对并撤销整篇排版…");
                var undone = await WordDocumentSynchronizer.UndoOutlineFormattingAsync(window.Document, history.Record, () => {
                    AssertOutlineUiThread(pane);
                    GetOperationCancellationToken().ThrowIfCancellationRequested();
                    if (!SameOutlineDocument(Application.ActiveDocument, history.Document) || !SameOutlineDocument(window.Document, history.Document))
                        throw new InvalidDataException("活动文档窗口已改变，已暂停整篇撤销。");
                });
                if (!undone.Item1) throw new InvalidDataException(undone.Item2);
                _outlineHistory.Remove(window.Hwnd);
                await RefreshReviewFromLive(pane, window.Document, window);
                SetStatus(pane, "已撤销本轮整篇排版；未修改其他格式属性和原文文字，未自动保存。");
                WriteStartupLog("已撤销整篇排版。");
            } catch (Exception exception) { WriteStartupLog("无法撤销整篇排版：" + exception);
                ShowFailure(pane, "无法撤销整篇排版：" + exception.Message);
                MessageBox.Show(exception.Message, "整篇撤销未完成", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            finally { SetBusy(pane, false, null); _ribbon?.Invalidate(); }
        }
    }
}
