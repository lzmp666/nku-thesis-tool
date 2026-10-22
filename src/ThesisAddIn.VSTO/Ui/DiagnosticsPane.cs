using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>只显示可检索的问题列表；分析、定位和修复命令位于 Word 功能区。</summary>
    public sealed class DiagnosticsPane : UserControl
    {
        private const string SearchPlaceholder = "搜索规则、期望值或实际值";
        private readonly TextBox _searchBox = new TextBox();
        private readonly ListBox _issueList = new ListBox();
        private readonly Label _status = new Label();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly List<EngineDiagnostic> _diagnostics = new List<EngineDiagnostic>();
        private readonly List<int> _visibleIndices = new List<int>();
        private int _selectedIndex = -1;
        private int _selectedOccurrenceIndex;
        private bool _isBusy;
        private bool _isShowingPlaceholder;
        private bool _updatingList;

        public event EventHandler DiagnosticSelected;
        public event EventHandler StateChanged;

        public DiagnosticsPane()
        {
            Dock = DockStyle.Fill;
            Font = new Font("Microsoft YaHei UI", 9F);

            _searchBox.Name = "IssueSearchBox";
            _searchBox.Dock = DockStyle.Fill;
            _searchBox.Margin = new Padding(0);
            _searchBox.AccessibleName = "搜索问题";
            _searchBox.AccessibleDescription = "按规则、期望格式或实际格式筛选问题。";
            _toolTip.SetToolTip(_searchBox, "可搜索规则名称、期望格式或实际格式");
            SetPlaceholder();
            _searchBox.GotFocus += OnSearchGotFocus;
            _searchBox.LostFocus += OnSearchLostFocus;
            _searchBox.TextChanged += OnSearchTextChanged;

            _issueList.Name = "IssueSearchResults";
            _issueList.Dock = DockStyle.Fill;
            _issueList.IntegralHeight = false;
            _issueList.HorizontalScrollbar = true;
            _issueList.AccessibleName = "格式问题列表";
            _issueList.SelectedIndexChanged += OnIssueSelected;

            _status.Dock = DockStyle.Fill;
            _status.Padding = new Padding(4, 8, 4, 2);
            _status.AutoEllipsis = true;
            _status.ForeColor = SystemColors.GrayText;
            _status.AccessibleName = "操作状态";

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(8)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));
            layout.Controls.Add(_searchBox, 0, 0);
            layout.Controls.Add(_issueList, 0, 1);
            layout.Controls.Add(_status, 0, 2);
            Controls.Add(layout);
        }

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
        public bool CanFixSelected => !_isBusy && IsFixable(SelectedDiagnostic);
        public bool HasFixableDiagnostics => !_isBusy && _diagnostics.Any(IsFixable);
        public bool CanNavigateIssues => !_isBusy && _visibleIndices.Count > 0;
        public bool CanNavigateOccurrences => !_isBusy && GetOccurrenceCount(SelectedDiagnostic) > 1;

        public void SetBusy(bool busy, string message)
        {
            InvokeOnUiThread(() =>
            {
                _isBusy = busy;
                _searchBox.Enabled = !busy;
                _issueList.Enabled = !busy;
                if (!string.IsNullOrEmpty(message)) SetStatusCore(message);
                RaiseStateChanged();
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
            _diagnostics.Clear();
            if (result?.Diagnostics != null)
                _diagnostics.AddRange(result.Diagnostics.Where(diagnostic => diagnostic != null));
            _selectedIndex = -1;
            _selectedOccurrenceIndex = 0;
            ClearSearch();
            RebuildIssueList();
            SetStatusCore(_diagnostics.Count == 0
                ? "未发现格式问题。结果仅覆盖已检查范围。"
                : "分析完成：" + _diagnostics.Count + " 个问题。可在 Word 功能区定位或修复。 ");
            RaiseStateChanged();
        }

        public void ShowFailure(string message)
        {
            InvokeOnUiThread(() =>
            {
                _diagnostics.Clear();
                _selectedIndex = -1;
                _selectedOccurrenceIndex = 0;
                RebuildIssueList();
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
            OnEvent(DiagnosticSelected);
        }

        private void SelectDiagnostic(int index, bool navigateToSource, string location = null)
        {
            if (index < 0 || index >= _diagnostics.Count) return;
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
                    _issueList.SelectedIndex = visibleIndex;
                    _issueList.TopIndex = Math.Max(0, visibleIndex);
                }
                finally { _updatingList = false; }
            }
            RaiseStateChanged();
            if (navigateToSource) OnEvent(DiagnosticSelected);
        }

        private void OnIssueSelected(object sender, EventArgs e)
        {
            if (_updatingList || _issueList.SelectedIndex < 0 || _issueList.SelectedIndex >= _visibleIndices.Count) return;
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
            RaiseStateChanged();
        }

        private void RebuildIssueList()
        {
            string query = _isShowingPlaceholder ? string.Empty : (_searchBox.Text ?? string.Empty).Trim();
            _visibleIndices.Clear();
            _issueList.BeginUpdate();
            try
            {
                _issueList.Items.Clear();
                for (int index = 0; index < _diagnostics.Count; index++)
                {
                    EngineDiagnostic diagnostic = _diagnostics[index];
                    if (!Matches(diagnostic, query)) continue;
                    _visibleIndices.Add(index);
                    _issueList.Items.Add(BuildIssueLabel(diagnostic));
                }
            }
            finally { _issueList.EndUpdate(); }

            int selectedVisibleIndex = _visibleIndices.IndexOf(_selectedIndex);
            if (selectedVisibleIndex >= 0)
            {
                _updatingList = true;
                try { _issueList.SelectedIndex = selectedVisibleIndex; }
                finally { _updatingList = false; }
            }
        }

        private static bool Matches(EngineDiagnostic diagnostic, string query)
        {
            if (string.IsNullOrEmpty(query)) return true;
            string searchable = string.Join(" ", new[]
            {
                diagnostic.RuleId, GetRuleName(diagnostic.RuleId), diagnostic.Expected, diagnostic.Actual,
                diagnostic.Message, diagnostic.FixLevel
            }.Where(value => !string.IsNullOrEmpty(value)));
            return searchable.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private static string BuildIssueLabel(EngineDiagnostic diagnostic)
        {
            int count = GetOccurrenceCount(diagnostic);
            return GetRuleName(diagnostic.RuleId) + " → " + (diagnostic.Expected ?? "（未提供）")
                + (count > 1 ? "（" + count + "处）" : string.Empty);
        }

        private static string GetRuleName(string ruleId)
        {
            switch (ruleId ?? string.Empty)
            {
                case "demo/east-asian-font": return "中文字体";
                case "demo/western-font": return "西文字体";
                case "demo/font-size": return "字号";
                default: return ruleId ?? "未命名问题";
            }
        }

        private void ClearSearch()
        {
            _updatingList = true;
            try
            {
                _searchBox.Text = string.Empty;
                _searchBox.ForeColor = SystemColors.WindowText;
                _isShowingPlaceholder = false;
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
