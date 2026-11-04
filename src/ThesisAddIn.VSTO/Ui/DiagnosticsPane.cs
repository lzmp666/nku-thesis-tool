using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ThesisAddIn.Protocol;
using ThesisAddIn.VSTO.Engine;

namespace ThesisAddIn.VSTO.Ui
{
    internal sealed class DiagnosticHoverEventArgs : EventArgs
    {
        internal EngineDiagnostic Diagnostic;
        internal string Location;
    }

    /// <summary>显示可检索的问题并保存当前文档的修复/撤销状态；原文建议显示为 Word 锚定批注。</summary>
    public sealed class DiagnosticsPane : UserControl
    {
        private const string SearchPlaceholder = "搜索类型、要求、实际或位置";
        private readonly TextBox _searchBox = new TextBox();
        private readonly Button _clearSearchButton = new Button();
        private readonly StableIssueListBox _issueList = new StableIssueListBox();
        private readonly Label _emptyResults = new Label();
        private readonly Label _summary = new Label();
        private readonly Label _ruleNotice = new Label();
        private readonly Label _status = new Label();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly Timer _navigationTimer = new Timer { Interval = 100 };
        private FlowLayoutPanel _batchBar;
        private Button _batchApply;
        private readonly Image _headerImage;
        private readonly List<EngineDiagnostic> _diagnostics = new List<EngineDiagnostic>();
        private readonly List<int> _visibleIndices = new List<int>();
        private readonly List<string> _visibleKeys = new List<string>();
        private int _selectedIndex = -1;
        private int _selectedOccurrenceIndex;
        private bool _isBusy;
        private bool _isShowingPlaceholder;
        private bool _updatingList;
        private bool _analysisSucceeded;
        private bool _selectionDuringBusy;
        private float _uiScale = 1F;
        private readonly List<Font> _scaledFonts = new List<Font>();

        internal readonly IssueReviewState Review;
        internal bool IsSuggestion { get; }
        internal void FocusIssueListForPreview()
        {
            if (IsSuggestion && !_issueList.IsDisposed && _issueList.CanFocus) _issueList.Focus();
        }
        internal void SyncReviewState() => OnReviewChanged(this, EventArgs.Empty);
        public event EventHandler RepairRequested;
        public event EventHandler UndoRequested;
        public event EventHandler ApplyCheckedRequested;
        public event EventHandler UndoCheckedRequested;
        public event EventHandler DiagnosticSelected;
        internal event EventHandler<DiagnosticHoverEventArgs> DiagnosticHoverChanged;
        public event EventHandler StateChanged;

        public DiagnosticsPane() : this(false, new IssueReviewState()) { }
        internal DiagnosticsPane(bool suggestions, IssueReviewState review)
        {
            IsSuggestion = suggestions;
            Review = review;
            Review.Changed += OnReviewChanged;
            Review.SelectionChanged += OnReviewSelectionChanged;
            Dock = DockStyle.Fill;
            DoubleBuffered = true;
            AutoScaleMode = AutoScaleMode.None;
            BackColor = Color.FromArgb(250, 247, 241);
            ForeColor = Color.FromArgb(44, 39, 42);
            Font = new Font("Microsoft YaHei UI", 9F);
            _navigationTimer.Tick += (sender, args) =>
            {
                _navigationTimer.Stop();
                if (!_isBusy && !IsDisposed && SelectedDiagnostic != null) OnEvent(DiagnosticSelected);
            };

            _summary.Dock = DockStyle.Top;
            _summary.Height = 27;
            _summary.Padding = new Padding(12, 5, 8, 3);
            _summary.Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold);
            _summary.ForeColor = Color.FromArgb(104, 31, 61);
            _summary.BackColor = Color.FromArgb(250, 247, 241);
            _summary.AutoEllipsis = true;
            _summary.TextAlign = ContentAlignment.MiddleLeft;
            _summary.Text = "格式问题";
            _summary.AccessibleName = "问题数量";
            _toolTip.SetToolTip(_summary,"当前为演示规则：支持正文中英文字体、字号与部分图表题注格式。图片布局、表格内部样式、流程图和公式尚未自动检查，不保证完整合规。");

            _ruleNotice.Dock = DockStyle.Fill;
            _ruleNotice.Text = "演示规则 · 用于验证功能，不代表正式合规结论";
            _ruleNotice.TextAlign = ContentAlignment.MiddleLeft;
            _ruleNotice.Padding = new Padding(8, 0, 5, 0);
            _ruleNotice.Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Bold);
            _ruleNotice.ForeColor = Color.FromArgb(116, 74, 28);
            _ruleNotice.BackColor = Color.FromArgb(250, 239, 216);
            _ruleNotice.AccessibleName = "规则状态：演示规则，不代表正式合规结论";
            _toolTip.SetToolTip(_ruleNotice,
                "当前 Word 插件连接的是 demo-0.1.0 演示规则，用于检查界面和修复流程；不是南开大学正式规则集，不能据此认定论文合规。");
            _ruleNotice.Paint += (sender, args) =>
            {
                using (var pen = new Pen(Color.FromArgb(224, 202, 165)))
                    args.Graphics.DrawRectangle(pen, 0, 0, _ruleNotice.Width - 1, _ruleNotice.Height - 1);
            };

            _searchBox.Name = "IssueSearchBox";
            _searchBox.Dock = DockStyle.Fill;
            _searchBox.BorderStyle = BorderStyle.None;
            _searchBox.Font = new Font("Microsoft YaHei UI", 9F);
            _searchBox.Margin = new Padding(0);
            _searchBox.AccessibleName = "搜索问题";
            _searchBox.AccessibleDescription = "按格式类别、期望值、实际值、位置、状态或来源筛选问题。";
            _toolTip.SetToolTip(_searchBox, "支持普通关键词，以及 类型:、期望:、实际:、位置:、状态:、来源: 字段搜索");
            SetPlaceholder();
            _searchBox.GotFocus += OnSearchGotFocus;
            _searchBox.LostFocus += OnSearchLostFocus;
            _searchBox.TextChanged += OnSearchTextChanged;
            _searchBox.KeyDown += OnSearchKeyDown;

            _clearSearchButton.Name = "ClearIssueSearch";
            _clearSearchButton.Dock = DockStyle.Fill;
            _clearSearchButton.Text = "×";
            _clearSearchButton.FlatStyle = FlatStyle.Flat;
            _clearSearchButton.FlatAppearance.BorderSize = 0;
            _clearSearchButton.BackColor = Color.Transparent;
            _clearSearchButton.ForeColor = Color.FromArgb(105, 119, 140);
            _clearSearchButton.Font = new Font("Microsoft YaHei UI", 13F, FontStyle.Regular);
            _clearSearchButton.Margin = new Padding(0);
            _clearSearchButton.TabStop = false;
            _clearSearchButton.Visible = false;
            _clearSearchButton.AccessibleName = "清除问题搜索";
            _clearSearchButton.Click += OnClearSearchClicked;

            _issueList.Name = "IssueSearchResults";
            _issueList.Dock = DockStyle.Fill;
            _issueList.BackColor = BackColor;
            _issueList.AccessibleName = "格式问题列表";
            _issueList.SelectedIndexChanged += OnIssueSelected;
            _issueList.CurrentItemInvoked += OnIssueSelected;
            _issueList.ItemPointerPressed += OnCardPointer;
            _issueList.ItemHoverChanged += OnCardHoverChanged;
            _issueList.DrawItem += OnDrawIssueItem;
            _issueList.MeasureItem += OnMeasureIssueItem;
            _issueList.KeyDown += OnIssueListKeyDown;

            _emptyResults.Dock = DockStyle.Fill;
            _emptyResults.Text = "没有匹配的问题\r\n调整关键词或清除搜索条件";
            _emptyResults.TextAlign = ContentAlignment.MiddleCenter;
            _emptyResults.ForeColor = Color.FromArgb(115, 126, 143);
            _emptyResults.BackColor = BackColor;
            _emptyResults.Visible = false;
            _emptyResults.AccessibleName = "没有匹配的问题";

            var resultsSurface = new Panel { Dock = DockStyle.Fill, BackColor = BackColor };
            resultsSurface.Controls.Add(_issueList);
            resultsSurface.Controls.Add(_emptyResults);

            _status.Dock = DockStyle.Fill;
            _status.Padding = new Padding(4, 8, 4, 2);
            _status.AutoEllipsis = true;
            _status.ForeColor = SystemColors.GrayText;
            _status.AccessibleName = "操作状态";

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(9, 6, 9, 8),
                BackColor = BackColor
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 29F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 0F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, IsSuggestion ? 0F : 38F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));

            _headerImage = null;
            var header = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(248, 242, 233),
                Margin = new Padding(0, 0, 0, 4),
                Padding = new Padding(0)
            };
            // Campus artwork lives in the Ribbon, leaving the index compact.
            header.Controls.Add(_summary);

            var searchSurface = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(9, 7, 8, 5),
                Margin = new Padding(0, 0, 0, 5)
            };
            searchSurface.Paint += (sender, args) =>
            {
                using (var pen = new Pen(Color.FromArgb(214, 222, 233)))
                    args.Graphics.DrawRectangle(pen, 0, 0, searchSurface.Width - 1, searchSurface.Height - 1);
            };
            var searchLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            searchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 22F));
            searchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            searchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 24F));
            var searchIcon = new Label
            {
                Dock = DockStyle.Fill,
                Text = "⌕",
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(105, 119, 140),
                Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Regular),
                Margin = new Padding(0)
            };
            searchLayout.Controls.Add(searchIcon, 0, 0);
            searchLayout.Controls.Add(_searchBox, 1, 0);
            searchLayout.Controls.Add(_clearSearchButton, 2, 0);
            searchSurface.Controls.Add(searchLayout);

            layout.Controls.Add(header, 0, 0);
            _ruleNotice.Visible = false;
            var batchBar = new FlowLayoutPanel { Dock = DockStyle.Fill, Margin = new Padding(0), WrapContents = false };
            _batchBar=batchBar;
            var all = new Button { Text = "全选", AutoSize = true, FlatStyle = FlatStyle.Flat };
            var none = new Button { Text = "清空", AutoSize = true, FlatStyle = FlatStyle.Flat };
            var apply = new Button { Text = "修复所选", AutoSize = true, FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(104,31,61), ForeColor = Color.White };
            _batchApply=apply;
            var close = new Button { Text = "退出", AutoSize = true, FlatStyle = FlatStyle.Flat };
            all.Click += (s, e) => Review.SelectAll(true);
            none.Click += (s, e) => Review.SelectAll(false);
            apply.Click += (s, e) => { if (!_isBusy) OnEvent(Review.BatchUndoMode ? UndoCheckedRequested : ApplyCheckedRequested); };
            close.Click += (s, e) => Review.SetBatch(false);
            batchBar.Controls.AddRange(new Control[] { all, none, apply, close });
            batchBar.Visible = false;
            Review.Changed += (s, e) => {
                if (batchBar.Visible != Review.BatchMode) batchBar.Visible = Review.BatchMode;
                float height = Review.BatchMode ? 34F * _uiScale : 0;
                if (layout.RowStyles[1].Height != height) layout.RowStyles[1].Height = height;
                apply.Text = Review.BatchUndoMode ? "撤销所选" : "修复所选";
                apply.Enabled = !_isBusy && Review.Checked.Count > 0;
            };
            layout.Controls.Add(batchBar, 0, 1);
            layout.Controls.Add(searchSurface, 0, 2);
            layout.Controls.Add(resultsSurface, 0, 3);
            layout.Controls.Add(_status, 0, 4);
            Controls.Add(layout);
        }

        internal int ConfigureForOfficeWindow(int windowHandle)
        {
            uint dpi = 96;
            try { dpi = GetDpiForWindow(new IntPtr(windowHandle)); }
            catch (EntryPointNotFoundException) { }
            _uiScale = Math.Max(1F, Math.Min(4F, dpi / 96F));
            var controls = new List<Control>();
            CollectControls(this, controls);
            var fonts = controls.Select(control => control.Font).ToArray();
            for (int i = 0; i < controls.Count; i++)
            {
                Control control = controls[i];
                var font = new Font(fonts[i].FontFamily, fonts[i].SizeInPoints * 96F / 72F * _uiScale,
                    fonts[i].Style, GraphicsUnit.Pixel);
                _scaledFonts.Add(font);
                control.Font = font;
                control.Padding = ScalePadding(control.Padding);
                control.Margin = ScalePadding(control.Margin);
                if (control is TableLayoutPanel table)
                {
                    foreach (RowStyle row in table.RowStyles)
                        if (row.SizeType == SizeType.Absolute) row.Height *= _uiScale;
                    foreach (ColumnStyle column in table.ColumnStyles)
                        if (column.SizeType == SizeType.Absolute) column.Width *= _uiScale;
                }
            }
            _summary.Height = (int)Math.Round(27F * _uiScale);
            return (int)Math.Round((IsSuggestion ? 265F : 300F) * _uiScale);
        }

        private Padding ScalePadding(Padding value) => new Padding(
            (int)Math.Round(value.Left * _uiScale), (int)Math.Round(value.Top * _uiScale),
            (int)Math.Round(value.Right * _uiScale), (int)Math.Round(value.Bottom * _uiScale));

        private static void CollectControls(Control control, List<Control> result)
        {
            result.Add(control);
            foreach (Control child in control.Controls) CollectControls(child, result);
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr window);

        private Font LogicalFont(float points, FontStyle style = FontStyle.Regular) =>
            new Font(Font.FontFamily, points * 96F / 72F, style, GraphicsUnit.Pixel);

        public EngineDiagnostic SelectedDiagnostic =>
            _selectedIndex >= 0 && _selectedIndex < _diagnostics.Count ? _diagnostics[_selectedIndex] : null;

        public string SelectedLocation
        {
            get
            {
                EngineDiagnostic diagnostic = SelectedDiagnostic;
                if (diagnostic == null) return null;
                if (diagnostic.Occurrences != null && diagnostic.Occurrences.Count > 0 &&
                    _selectedOccurrenceIndex >= 0 && _selectedOccurrenceIndex < diagnostic.Occurrences.Count)
                    return diagnostic.Occurrences[_selectedOccurrenceIndex].Location;
                return diagnostic.Location;
            }
        }

        public IList<EngineDiagnostic> Diagnostics => _diagnostics;
        public bool IsBusy => _isBusy;
        public bool CanFixSelected => !_isBusy && Review.CanRepair(SelectedDiagnostic);
        public bool CanResolveSelected => !_isBusy && Review.IsRepaired(SelectedDiagnostic);
        public bool HasFixableDiagnostics => !_isBusy && _diagnostics.Any(Review.CanRepair);
        public bool HasUndoableDiagnostics => !_isBusy && _diagnostics.Any(Review.IsRepaired);
        public bool CanNavigateIssues => !_isBusy && _visibleIndices.Count > 0;
        public bool CanNavigateOccurrences => !_isBusy && GetOccurrenceCount(SelectedDiagnostic) > 1;

        public void SetBusy(bool busy, string message)
        {
            InvokeOnUiThread(() =>
            {
                _isBusy = busy;
                foreach (Control command in _batchBar.Controls) command.Enabled=!busy;
                _batchApply.Enabled=!busy && Review.Checked.Count>0;
                if (busy) _navigationTimer.Stop();
                // Keep search/scroll/selection responsive while writes and rechecks
                // are in progress. Only commands that mutate Word are blocked.
                _searchBox.Enabled = true;
                _issueList.Enabled = true;
                _issueList.Invalidate();
                if (!string.IsNullOrEmpty(message)) SetStatusCore(message);
                RaiseStateChanged();
                if (!busy && _selectionDuringBusy) { _selectionDuringBusy=false; QueueSourceNavigation(); }
            });
        }

        public void SetStatus(string message)
        {
            InvokeOnUiThread(() => SetStatusCore(message));
        }

        public void ShowAnalysis(EngineAnalysisResult result)
        {
            InvokeOnUiThread(() => ShowAnalysisCore(result));
        }

        private void ShowAnalysisCore(EngineAnalysisResult result)
        {
            _navigationTimer.Stop();
            _analysisSucceeded = true;
            Review.Refresh(result?.Diagnostics);
            string baseStatus = _diagnostics.Count == 0
                ? "未发现已覆盖范围内的格式问题；未覆盖项不会被判为合规。"
                : "点击问题定位原文；批注中的“解决”可修复，撤销请选中问题后点击功能区“撤销当前修复”。";
            string[] unsupported = result?.UnsupportedScopes?.Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(DescribeUnsupportedScope).Distinct(StringComparer.CurrentCulture).ToArray()
                ?? Array.Empty<string>();
            string unsupportedTooltip = null;
            if (unsupported.Length > 0)
            {
                string visible = string.Join("、", unsupported.Take(3));
                if (unsupported.Length > 3) visible += "等" + unsupported.Length + "类";
                baseStatus += Environment.NewLine + "未自动检查：" + visible + "（悬停查看范围）";
                unsupportedTooltip = "未自动检查的范围：" + Environment.NewLine +
                    string.Join(Environment.NewLine, unsupported.Select(value => "· " + value));
            }
            SetStatusCore(baseStatus);
            if (!string.IsNullOrEmpty(unsupportedTooltip)) _toolTip.SetToolTip(_status, unsupportedTooltip);
            RaiseStateChanged();
        }

        public void ShowFailure(string message)
        {
            InvokeOnUiThread(() =>
            {
                _analysisSucceeded = false;
                _diagnostics.Clear();
                _selectedIndex = -1;
                _selectedOccurrenceIndex = 0;
                RebuildIssueList();
                _summary.Text = "格式问题";
                SetStatusCore("分析或修复未完成：" + (string.IsNullOrWhiteSpace(message) ? "未知错误。" : message.Trim()));
                RaiseStateChanged();
            });
        }

        /// <summary>
        /// Engine protocol calls can resume on a worker thread. Keep every WinForms
        /// control access on the Word UI thread and surface exceptions to the caller.
        /// </summary>
        public T InvokeOnUiThread<T>(Func<T> action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            if (IsDisposed || Disposing || !IsHandleCreated)
                throw new InvalidOperationException("问题检索列表已关闭，无法更新界面。");
            return InvokeRequired ? (T)Invoke(action) : action();
        }

        public void InvokeOnUiThread(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            InvokeOnUiThread(() =>
            {
                action();
                return true;
            });
        }

        /// <summary>Selects the unique grouped issue described by a Word comment.</summary>
        public bool SelectByComment(string ruleId, string expected, string actual, string location, bool navigateToSource)
        {
            if (string.IsNullOrWhiteSpace(ruleId)) return false;
            int matchIndex = CommentDiagnosticMatcher.FindUniqueMatchIndex(
                _diagnostics, ruleId, expected, actual, location);
            if (matchIndex < 0) return false;
            SelectDiagnostic(matchIndex, navigateToSource, location);
            return true;
        }

        public bool SelectByOccurrenceLocation(string location, bool navigateToSource)
        {
            if (string.IsNullOrWhiteSpace(location)) return false;
            for (int index = 0; index < _diagnostics.Count; index++)
            {
                EngineDiagnostic diagnostic = _diagnostics[index];
                if (diagnostic?.Occurrences == null) continue;
                if (diagnostic.Occurrences.Any(occurrence => occurrence != null &&
                    string.Equals(occurrence.Location, location, StringComparison.Ordinal)))
                {
                    SelectDiagnostic(index, navigateToSource, location);
                    return true;
                }
            }
            return false;
        }

        public void NavigateIssue(int direction)
        {
            if (_isBusy || direction == 0 || _visibleIndices.Count == 0) return;
            int currentVisibleIndex = _visibleIndices.IndexOf(_selectedIndex);
            int nextVisibleIndex = currentVisibleIndex < 0
                ? (direction > 0 ? 0 : _visibleIndices.Count - 1)
                : (currentVisibleIndex + direction + _visibleIndices.Count) % _visibleIndices.Count;
            SelectDiagnostic(_visibleIndices[nextVisibleIndex], navigateToSource: true);
        }

        public void NavigateOccurrence(int direction)
        {
            if (_isBusy || direction == 0 || SelectedDiagnostic == null) return;
            int count = GetOccurrenceCount(SelectedDiagnostic);
            if (count <= 1) return;
            _selectedOccurrenceIndex = (_selectedOccurrenceIndex + direction + count) % count;
            RaiseStateChanged();
            QueueSourceNavigation();
        }

        private void SelectDiagnostic(int index, bool navigateToSource, string location = null)
        {
            if (index < 0 || index >= _diagnostics.Count) return;
            int priorIndex = _selectedIndex;
            int priorOccurrence = _selectedOccurrenceIndex;
            _selectedIndex = index;
            EngineDiagnostic selected = _diagnostics[index];
            _selectedOccurrenceIndex = 0;
            if (!string.IsNullOrWhiteSpace(location) && selected.Occurrences != null)
            {
                for (int i = 0; i < selected.Occurrences.Count; i++)
                {
                    if (string.Equals(selected.Occurrences[i].Location, location, StringComparison.Ordinal))
                    {
                        _selectedOccurrenceIndex = i;
                        break;
                    }
                }
            }

            int visibleIndex = _visibleIndices.IndexOf(index);
            if (visibleIndex >= 0)
            {
                _updatingList = true;
                try
                {
                    int offset = _issueList.ScrollOffset;
                    if (_issueList.SelectedIndex != visibleIndex)
                        _issueList.SelectedIndex = visibleIndex;
                    // Body/comment hover must not scroll the searchable index either.
                    _issueList.RestoreScrollOffset(offset);
                }
                finally { _updatingList = false; }
            }
            if (priorIndex != _selectedIndex || priorOccurrence != _selectedOccurrenceIndex)
            {
                if (visibleIndex >= 0) _issueList.Invalidate(_issueList.GetItemRectangle(visibleIndex));
                RaiseStateChanged();
            }
            Review.Select(selected, location ?? SelectedLocation);
            if (navigateToSource) QueueSourceNavigation();
        }

        // Let the native ListBox finish selection/painting before entering Word COM.
        // Rapid clicks replace pending work; only the latest selected issue is located.
        private void QueueSourceNavigation()
        {
            _navigationTimer.Stop();
            _navigationTimer.Start();
        }

        private void OnIssueSelected(object sender, EventArgs e)
        {
            if (_updatingList || _issueList.SelectedIndex < 0 || _issueList.SelectedIndex >= _visibleIndices.Count) return;
            if (_isBusy) {
                SelectDiagnostic(_visibleIndices[_issueList.SelectedIndex], navigateToSource:false);
                _selectionDuringBusy=true; return;
            }
            SelectDiagnostic(_visibleIndices[_issueList.SelectedIndex], navigateToSource: true);
        }

        private void OnSearchGotFocus(object sender, EventArgs e)
        {
            if (!_isShowingPlaceholder) return;
            _isShowingPlaceholder = false;
            _searchBox.Text = string.Empty;
            _searchBox.ForeColor = SystemColors.WindowText;
        }

        private void OnSearchLostFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_searchBox.Text)) SetPlaceholder();
        }

        private void OnSearchTextChanged(object sender, EventArgs e)
        {
            if (_isShowingPlaceholder || _updatingList) return;
            int priorSelection = _selectedIndex;
            RebuildIssueList();
            if (!_visibleIndices.Contains(priorSelection))
            {
                _selectedIndex = -1;
                _selectedOccurrenceIndex = 0;
            }
            else
            {
                int visibleIndex = _visibleIndices.IndexOf(priorSelection);
                _updatingList = true;
                try { _issueList.SelectedIndex = visibleIndex; }
                finally { _updatingList = false; }
            }
            _clearSearchButton.Visible = !string.IsNullOrWhiteSpace(_searchBox.Text);
            RaiseStateChanged();
        }

        private void OnClearSearchClicked(object sender, EventArgs e)
        {
            ClearSearch();
            RebuildIssueList();
            RaiseStateChanged();
        }

        private void OnSearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                ClearSearch();
                RebuildIssueList();
                RaiseStateChanged();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter && _selectedIndex >= 0)
            {
                SelectDiagnostic(_selectedIndex, navigateToSource: true);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Down && _visibleIndices.Count > 0)
            {
                _issueList.Focus();
                if (_issueList.SelectedIndex < 0) _issueList.SelectedIndex = 0;
                e.Handled = true;
            }
        }

        private void OnIssueListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && _issueList.SelectedIndex >= 0 &&
                _issueList.SelectedIndex < _visibleIndices.Count)
            {
                SelectDiagnostic(_visibleIndices[_issueList.SelectedIndex], navigateToSource: true);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                ClearSearch();
                RebuildIssueList();
                _searchBox.Focus();
                RaiseStateChanged();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void RebuildIssueList()
        {
            string query = _isShowingPlaceholder ? string.Empty : (_searchBox.Text ?? string.Empty).Trim();
            var next = Enumerable.Range(0,_diagnostics.Count).Where(index => Matches(_diagnostics[index],query)).ToArray();
            string[] keys = next.Select(index => IssueReviewState.Key(_diagnostics[index])).ToArray();
            int offset = _issueList.ScrollOffset;
            bool unchanged = _visibleKeys.SequenceEqual(keys);
            _updatingList = true;
            try {
                _visibleIndices.Clear(); _visibleIndices.AddRange(next);
                if (!unchanged) {
                    _visibleKeys.Clear(); _visibleKeys.AddRange(keys);
                    _issueList.BeginUpdate();
                    try {
                        _issueList.Items.Clear();
                        foreach (int index in next) _issueList.Items.Add(BuildIssueLabel(_diagnostics[index]));
                    } finally { _issueList.EndUpdate(); }
                } else {
                    for (int i=0;i<next.Length;i++) _issueList.Items[i]=BuildIssueLabel(_diagnostics[next[i]]);
                    _issueList.RefreshLayout();
                }
                _issueList.RestoreScrollOffset(offset);
            } finally { _updatingList=false; }

            int selectedVisibleIndex = _visibleIndices.IndexOf(_selectedIndex);
            if (selectedVisibleIndex >= 0)
            {
                _updatingList = true;
                try { _issueList.SelectedIndex = selectedVisibleIndex; }
                finally { _updatingList = false; }
            }

            int visibleOccurrences = _visibleIndices.Sum(index => GetOccurrenceCount(_diagnostics[index]));
            int totalOccurrences = _diagnostics.Sum(GetOccurrenceCount);
            _summary.Text = string.IsNullOrEmpty(query)
                ? "待处理 " + _diagnostics.Count(x=>!Review.IsRepaired(x)) + " 项 · 已修复 " + Review.Repaired.Count + " 项"
                : "检索结果  " + _visibleIndices.Count + " 项 · " + visibleOccurrences + " 处（共 "
                    + _diagnostics.Count + " 项）";
            _emptyResults.Visible = _analysisSucceeded && _visibleIndices.Count == 0;
            _emptyResults.Text = query.Length > 0
                ? "没有匹配的问题\r\n调整关键词或清除搜索条件"
                : "未发现已覆盖范围内的问题";
            _clearSearchButton.Visible = query.Length > 0;
        }

        private static bool Matches(EngineDiagnostic diagnostic, string query)
        {
            if (string.IsNullOrEmpty(query)) return true;
            string type = DescribeRule(diagnostic.RuleId);
            string location = GetLocations(diagnostic);
            string status = DescribeFixLevel(diagnostic.FixLevel) + " " + DescribeSeverity(diagnostic.Severity);
            string source = GetSource(diagnostic);
            string[] terms = query.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            foreach (string term in terms)
            {
                int separator = term.IndexOfAny(new[] { ':', '：' });
                string field = separator < 0 ? string.Empty : term.Substring(0, separator).Trim();
                string value = separator < 0 ? term : term.Substring(separator + 1).Trim();
                if (value.Length == 0) return false;
                string haystack;
                switch (field)
                {
                    case "类型": haystack = type; break;
                    case "要求": case "期望": haystack = diagnostic.Expected; break;
                    case "实际": haystack = GetActualValues(diagnostic); break;
                    case "位置": haystack = location; break;
                    case "状态": haystack = status; break;
                    case "来源": haystack = source; break;
                    default:
                        haystack = string.Join(" ", new[] { type, diagnostic.Expected,
                            GetActualValues(diagnostic), location, status, source, diagnostic.Message });
                        value = term;
                        break;
                }
                if ((haystack ?? string.Empty).IndexOf(value, StringComparison.CurrentCultureIgnoreCase) < 0)
                    return false;
            }
            return true;
        }

        private static string BuildIssueLabel(EngineDiagnostic diagnostic)
        {
            int count = GetOccurrenceCount(diagnostic);
            return DescribeRule(diagnostic.RuleId) + "　期望 " + (diagnostic.Expected ?? "（未提供）")
                + (count > 1 ? "　" + count + "处" : string.Empty);
        }

        public static string DescribeRule(string ruleId)
        {
            switch (ruleId ?? string.Empty)
            {
                case "demo/east-asian-font": return "中文字体";
                case "NKU-REQ-042": return "中文字体";
                case "demo/western-font": return "西文字体";
                case "NKU-REQ-043": return "西文字体";
                case "demo/font-size": return "字号";
                case "NKU-REQ-044": return "字号";
                case "demo/caption-east-asian-font": return "图表题注字体";
                case "demo/caption-font-size": return "图表题注字号";
                case "demo/caption-alignment": return "图表题注对齐";
                case "demo/caption-line-spacing": return "图表题注行距";
                case "demo/caption-space-before": return "图表题注段前距";
                case "demo/caption-space-after": return "图表题注段后距";
                case "demo/caption-placement": return "图表题注位置";
                case "demo/caption-number-gap": return "题注编号间距";
                case "demo/picture-fit-width": return "图片超出版心";
                case "demo/table-fit-width": return "表格超出版心";
                default:
                    if (!string.IsNullOrWhiteSpace(ruleId) && ruleId.IndexOf("caption", StringComparison.OrdinalIgnoreCase) >= 0)
                        return "图表题注";
                    return "其他格式";
            }
        }

        private void OnMeasureIssueItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _visibleIndices.Count) { e.ItemHeight = 90; return; }
            EngineDiagnostic issue = _diagnostics[_visibleIndices[e.Index]];
            int width = Math.Max(140, (int)(_issueList.ContentWidth / _uiScale) - 12);
            using (var font = LogicalFont(8.8F)) {
                int height = 29;
                height += MeasureCardLineHeight(e.Graphics, issue.Expected, 12, width-12, font)+2;
                height += MeasureCardLineHeight(e.Graphics, ActualText(issue), 12, width-12, font)+2;
                if (!IsSuggestion) height += MeasureCardLineHeight(e.Graphics,
                    GetCompactLocation(_visibleIndices[e.Index], issue), 12, width-12, font)+2;
                if (IsSuggestion) height += 30;
                e.ItemHeight = (int)Math.Ceiling(height * _uiScale);
            }
        }
        private string ActualText(EngineDiagnostic issue) => Review.IsRepaired(issue)
            ? DisplayValue(issue.Expected) : DisplayValue(GetActualValues(issue));
        private static string DisplayValue(string value)
        {
            switch (value) { case "center": return "居中"; case "left": return "左对齐";
                case "right": return "右对齐"; case "single": return "单倍行距";
                case "multiple:360": return "1.5 倍行距"; case "(unknown)": return "未知，需核对";
                default: return value ?? "未知"; }
        }
        private void OnDrawIssueItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _visibleIndices.Count) return;
            GraphicsState saved = e.Graphics.Save();
            try {
                e.Graphics.ScaleTransform(_uiScale, _uiScale);
                Rectangle bounds = new Rectangle((int)(e.Bounds.X/_uiScale), (int)(e.Bounds.Y/_uiScale),
                    (int)(e.Bounds.Width/_uiScale), (int)(e.Bounds.Height/_uiScale));
                DrawIssueCard(e.Graphics, bounds, e.Index, (e.State & DrawItemState.Selected) != 0);
            } finally { e.Graphics.Restore(saved); }
        }
        private void DrawIssueCard(Graphics graphics, Rectangle bounds, int index, bool selected)
        {
            EngineDiagnostic issue = _diagnostics[_visibleIndices[index]];
            Rectangle card = Rectangle.Inflate(bounds, -3,-3);
            Color purple = Color.FromArgb(104,31,61);
            using (GraphicsPath path = RoundedRectangle(card,6))
            using (var fill = new SolidBrush(selected ? Color.FromArgb(255,249,233) : Color.White))
            using (var border = new Pen(selected ? purple : Color.FromArgb(225,218,211), selected?1.5F:1F)) {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.FillPath(fill,path); graphics.DrawPath(border,path);
            }
            using (var title = LogicalFont(9F,FontStyle.Bold))
            using (var body = LogicalFont(8.8F))
            using (var ink = new SolidBrush(ForeColor))
            using (var accent = new SolidBrush(purple)) {
                float y = card.Top+7;
                bool check = Review.Checked.Contains(IssueReviewState.Key(issue));
                string lead = Review.BatchMode ? (Review.CanSelectForBatch(issue) ? (check ? "☑ " : "☐ ") : "— ") : "";
                string state = Review.IsRepaired(issue) ? " · 已修复" : "";
                graphics.DrawString(lead+DescribeRule(issue.RuleId)+state,title,accent,card.Left+9,y);
                y += 22;
                DrawCardLine(graphics,"要求",DisplayValue(issue.Expected),card.Left+9,ref y,card.Right-9,body,ink);
                DrawCardLine(graphics,"实际",ActualText(issue),card.Left+9,ref y,card.Right-9,body,ink);
                if (!IsSuggestion)
                    DrawCardLine(graphics,"位置",GetCompactLocation(_visibleIndices[index],issue),
                        card.Left+9,ref y,card.Right-9,body,ink);
                if (IsSuggestion)
                {
                    int buttonY = card.Bottom-28;
                    DrawAction(graphics,new Rectangle(card.Left+9,buttonY,66,23),"解决",!_isBusy && Review.CanRepair(issue));
                    DrawAction(graphics,new Rectangle(card.Left+83,buttonY,66,23),"撤销",!_isBusy && Review.IsRepaired(issue));
                }
            }
        }
        private void DrawAction(Graphics graphics, Rectangle rect, string label, bool enabled)
        {
            using (var fill = new SolidBrush(enabled?Color.FromArgb(104,31,61):Color.FromArgb(237,233,229)))
            using (var ink = new SolidBrush(enabled?Color.White:Color.Gray))
            using (var font = LogicalFont(8.5F)) {
                graphics.FillRectangle(fill,rect);
                graphics.DrawString(label,font,ink,rect,new StringFormat { Alignment=StringAlignment.Center,LineAlignment=StringAlignment.Center });
            }
        }
        private void OnCardPointer(object sender, CardPointerEventArgs e)
        {
            if (_isBusy || e.Index<0 || e.Index>=_visibleIndices.Count) return;
            int index = _visibleIndices[e.Index];
            EngineDiagnostic issue = _diagnostics[index];
            Rectangle bounds = _issueList.GetItemRectangle(e.Index);
            float x=e.X/_uiScale, y=(e.Y-bounds.Top)/_uiScale, height=bounds.Height/_uiScale;
            if (Review.BatchMode && x<36) {
                SelectDiagnostic(index,false); Review.Toggle(issue); e.Handled=true; return;
            }
            if (IsSuggestion && y>=height-33) {
                SelectDiagnostic(index,false);
                if (x>=12 && x<=80 && Review.CanRepair(issue)) OnEvent(RepairRequested);
                else if (x>=86 && x<=155 && Review.IsRepaired(issue)) OnEvent(UndoRequested);
                e.Handled=true;
            }
        }
        private void OnCardHoverChanged(object sender, CardHoverEventArgs e)
        {
            if (!IsSuggestion) return;
            if (_isBusy)
            {
                DiagnosticHoverChanged?.Invoke(this, new DiagnosticHoverEventArgs());
                return;
            }
            if (e.Index < 0 || e.Index >= _visibleIndices.Count)
            {
                DiagnosticHoverChanged?.Invoke(this, new DiagnosticHoverEventArgs());
                return;
            }
            EngineDiagnostic issue = _diagnostics[_visibleIndices[e.Index]];
            if (!Review.CanRepair(issue))
            {
                DiagnosticHoverChanged?.Invoke(this, new DiagnosticHoverEventArgs());
                return;
            }
            string location = issue.Occurrences != null && issue.Occurrences.Count > 0
                ? issue.Occurrences[0]?.Location : issue.Location;
            DiagnosticHoverChanged?.Invoke(this, new DiagnosticHoverEventArgs
            { Diagnostic = issue, Location = location });
        }
        private void OnReviewChanged(object sender, EventArgs e)
        {
            _analysisSucceeded=true;
            if (_diagnostics.SequenceEqual(Review.Rows)) {
                _summary.Text="待处理 "+_diagnostics.Count(x=>!Review.IsRepaired(x))+" 项 · 已修复 "+Review.Repaired.Count+" 项";
                _issueList.Invalidate(); RaiseStateChanged(); return;
            }
            _diagnostics.Clear(); _diagnostics.AddRange(Review.Rows);
            _selectedIndex = _diagnostics.FindIndex(x=>IssueReviewState.Key(x)==Review.SelectedKey);
            RebuildIssueList(); RaiseStateChanged();
        }
        private void OnReviewSelectionChanged(object sender, EventArgs e)
        {
            int index=_diagnostics.FindIndex(x=>IssueReviewState.Key(x)==Review.SelectedKey);
            if(index>=0) {
                SelectDiagnostic(index,false,Review.SelectedLocation);
                // Reveal the corresponding actionable card, never scroll the search index.
                if(IsSuggestion) _issueList.RevealSelection();
            }
        }
        internal void BeginBatch() { Review.SetBatch(true); }
        internal void BeginUndoBatch() { Review.SetBatchMode(true,true); }
        internal EngineDiagnostic[] CheckedDiagnostics => _diagnostics.Where(x=>Review.Checked.Contains(IssueReviewState.Key(x)) && Review.CanRepair(x)).ToArray();
        internal EngineDiagnostic[] CheckedUndoDiagnostics => _diagnostics.Where(x=>Review.Checked.Contains(IssueReviewState.Key(x)) && Review.IsRepaired(x)).ToArray();

        private static void DrawCardLine(Graphics graphics, string label, string value, int left,
            ref float y, int right, Font font, Brush valueBrush)
        {
            using (var labelFont = new Font(font.FontFamily, font.Size, FontStyle.Bold, font.Unit))
            using (var labelBrush = new SolidBrush(Color.FromArgb(112, 99, 102)))
            {
                const int labelColumnWidth = 43;
                int textWidth = Math.Max(80, right - left - labelColumnWidth);
                int rowHeight = MeasureCardLineHeight(graphics, value, left, right, font);
                graphics.DrawString(label, labelFont, labelBrush, left, y);
                var text = new RectangleF(left + labelColumnWidth, y, textWidth, rowHeight);
                graphics.DrawString(value ?? "—", font, valueBrush, text);
                y += rowHeight + 2;
            }
        }

        private static int MeasureCardLineHeight(Graphics graphics, string value, int left, int right, Font font)
        {
            const int labelColumnWidth = 43;
            int textWidth = Math.Max(80, right - left - labelColumnWidth);
            SizeF size = graphics.MeasureString(string.IsNullOrWhiteSpace(value) ? "—" : value,
                font, textWidth);
            return Math.Max((int)Math.Ceiling(font.GetHeight(graphics)) + 3, (int)Math.Ceiling(size.Height) + 2);
        }

        private static Image LoadHeaderImage()
        {
            using (Stream stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("ThesisAddIn.VSTO.Assets.nku-campus-panorama-v2.png"))
            {
                if (stream == null) return null;
                using (Image source = Image.FromStream(stream)) return new Bitmap(source);
            }
        }

        private void PaintHeader(Graphics graphics, Rectangle bounds)
        {
            if (_headerImage == null || bounds.Width <= 0 || bounds.Height <= 0) return;
            int sourceHeight = Math.Min(_headerImage.Height,
                Math.Max(1, (int)Math.Ceiling(_headerImage.Width * bounds.Height / (double)bounds.Width)));
            int sourceTop = Math.Max(0, (_headerImage.Height - sourceHeight) / 2);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.DrawImage(_headerImage, bounds,
                new Rectangle(0, sourceTop, _headerImage.Width, sourceHeight), GraphicsUnit.Pixel);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Review.Changed -= OnReviewChanged;
                Review.SelectionChanged -= OnReviewSelectionChanged;
                _navigationTimer.Stop();
                _navigationTimer.Dispose();
                _headerImage?.Dispose();
                _toolTip.Dispose();
            }
            base.Dispose(disposing);
            if (disposing) foreach (Font font in _scaledFonts) font.Dispose();
        }

        private static GraphicsPath RoundedRectangle(Rectangle bounds, int radius)
        {
            int diameter = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }

        private static string GetActualValues(EngineDiagnostic diagnostic)
        {
            if (diagnostic?.Occurrences == null || diagnostic.Occurrences.Count == 0)
                return diagnostic?.Actual ?? "未知";
            string[] actuals = diagnostic.Occurrences.Select(item => item?.Actual)
                .Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.CurrentCulture).ToArray();
            return actuals.Length == 0 ? diagnostic?.Actual ?? "未知" : string.Join("；", actuals);
        }

        private static string GetLocations(EngineDiagnostic diagnostic)
        {
            string[] locations = diagnostic?.Occurrences?.Select(item => item?.Location)
                .Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal).ToArray();
            if (locations == null || locations.Length == 0) locations = new[] { diagnostic?.Location };
            return string.Join("、", locations.Select(location =>
            {
                DocumentLocation parsed;
                string error;
                return DocumentLocator.TryParse(location, out parsed, out error)
                    ? DocumentLocator.Describe(parsed) : location;
            }));
        }

        private string GetCompactLocation(int diagnosticIndex, EngineDiagnostic diagnostic)
        {
            int count = GetOccurrenceCount(diagnostic);
            int occurrenceIndex = diagnosticIndex == _selectedIndex
                ? Math.Max(0, Math.Min(_selectedOccurrenceIndex, count - 1)) : 0;
            string location = diagnostic?.Occurrences != null && diagnostic.Occurrences.Count > occurrenceIndex
                ? diagnostic.Occurrences[occurrenceIndex]?.Location : diagnostic?.Location;
            DocumentLocation parsed;
            string error;
            string description = DocumentLocator.TryParse(location, out parsed, out error)
                ? DocumentLocator.Describe(parsed) : location ?? "未知位置";
            if (count <= 1) return description;
            return description + "（" + (occurrenceIndex + 1) + "/" + count + "）";
        }

        private static string GetSource(EngineDiagnostic diagnostic)
        {
            string message = diagnostic?.Message ?? string.Empty;
            int separator = message.IndexOf('：');
            if (separator > 0 && (diagnostic?.RuleId ?? string.Empty).StartsWith("NKU-REQ-", StringComparison.Ordinal))
                return message.Substring(0, separator);
            // The single banner at the top identifies demo mode. Repeating this on every
            // card obscures the expected/actual values without adding a real source.
            return null;
        }

        private static string BuildRecommendation(EngineDiagnostic diagnostic)
        {
            if (IsFixable(diagnostic))
                return "可自动修复为「" + (diagnostic.Expected ?? "") + "」";
            return "需人工核对，插件不会自动修改";
        }

        private static string DescribeUnsupportedScope(string value)
        {
            string scope = value ?? string.Empty;
            if (scope.IndexOf("tables", StringComparison.OrdinalIgnoreCase) >= 0) return "表格单元格专属格式";
            if (scope.IndexOf("caption", StringComparison.OrdinalIgnoreCase) >= 0) return "图题/表题格式";
            if (scope.IndexOf("images", StringComparison.OrdinalIgnoreCase) >= 0) return "图片清晰度与版式";
            if (scope.IndexOf("equation", StringComparison.OrdinalIgnoreCase) >= 0) return "公式内容及排版例外";
            if (scope.IndexOf("footnote", StringComparison.OrdinalIgnoreCase) >= 0) return "脚注格式";
            if (scope.IndexOf("endnote", StringComparison.OrdinalIgnoreCase) >= 0) return "尾注格式";
            if (scope.IndexOf("header", StringComparison.OrdinalIgnoreCase) >= 0) return "页眉页脚";
            if (scope.IndexOf("textbox", StringComparison.OrdinalIgnoreCase) >= 0) return "文本框内容";
            if (scope.IndexOf("tracked changes", StringComparison.OrdinalIgnoreCase) >= 0) return "修订内容";
            return scope;
        }

        private void ClearSearch()
        {
            _updatingList = true;
            try
            {
                _searchBox.Text = string.Empty;
                _searchBox.ForeColor = SystemColors.WindowText;
                _isShowingPlaceholder = false;
                if (!_searchBox.Focused) SetPlaceholder();
            }
            finally { _updatingList = false; }
        }

        private void SetPlaceholder()
        {
            _isShowingPlaceholder = true;
            _searchBox.ForeColor = SystemColors.GrayText;
            _searchBox.Text = SearchPlaceholder;
        }

        private void SetStatusCore(string message)
        {
            _status.Text = message ?? string.Empty;
            _toolTip.SetToolTip(_status, _status.Text);
        }

        private void RaiseStateChanged() => OnEvent(StateChanged);

        private void OnEvent(EventHandler handler)
        {
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private static int GetOccurrenceCount(EngineDiagnostic diagnostic) =>
            diagnostic?.Occurrences != null && diagnostic.Occurrences.Count > 0 ? diagnostic.Occurrences.Count : 1;

        private static bool IsFixable(EngineDiagnostic diagnostic) =>
            diagnostic != null && string.Equals(diagnostic.FixLevel, "auto", StringComparison.OrdinalIgnoreCase);

        public static string DescribeSeverity(string severity)
        {
            switch ((severity ?? string.Empty).ToLowerInvariant())
            {
                case "blocking": return "阻塞";
                case "warning": return "警告";
                case "info": return "提示";
                default: return string.IsNullOrEmpty(severity) ? "未知" : severity;
            }
        }

        public static string DescribeFixLevel(string fixLevel)
        {
            switch ((fixLevel ?? string.Empty).ToLowerInvariant())
            {
                case "auto": return "可自动";
                case "suggest": return "需确认";
                case "manual": return "仅人工";
                default: return string.IsNullOrEmpty(fixLevel) ? "未知" : fixLevel;
            }
        }
    }
}
