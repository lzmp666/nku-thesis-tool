using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using ThesisAddIn.Protocol;
using ThesisAddIn.VSTO.Engine;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>只显示可检索的问题列表；分析、定位和修复命令位于 Word 功能区。</summary>
    public sealed class DiagnosticsPane : UserControl
    {
        private const string SearchPlaceholder = "搜索问题（例：类型:西文字体 期望:Times New Roman）";
        private readonly TextBox _searchBox = new TextBox();
        private readonly Button _clearSearchButton = new Button();
        private readonly ListBox _issueList = new ListBox();
        private readonly Label _emptyResults = new Label();
        private readonly Label _summary = new Label();
        private readonly Label _ruleNotice = new Label();
        private readonly Label _status = new Label();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly List<EngineDiagnostic> _diagnostics = new List<EngineDiagnostic>();
        private readonly List<int> _visibleIndices = new List<int>();
        private int _selectedIndex = -1;
        private int _selectedOccurrenceIndex;
        private bool _isBusy;
        private bool _isShowingPlaceholder;
        private bool _updatingList;
        private bool _analysisSucceeded;

        public event EventHandler DiagnosticSelected;
        public event EventHandler StateChanged;

        public DiagnosticsPane()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(246, 248, 251);
            ForeColor = Color.FromArgb(35, 43, 56);
            Font = new Font("Microsoft YaHei UI", 9F);

            _summary.Dock = DockStyle.Fill;
            _summary.Padding = new Padding(4, 7, 4, 2);
            _summary.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
            _summary.Text = "格式问题";
            _summary.AccessibleName = "问题数量";

            _ruleNotice.Dock = DockStyle.Fill;
            _ruleNotice.Text = "演示规则 · 用于验证功能，不代表正式合规结论";
            _ruleNotice.TextAlign = ContentAlignment.MiddleLeft;
            _ruleNotice.Padding = new Padding(8, 0, 5, 0);
            _ruleNotice.Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Bold);
            _ruleNotice.ForeColor = Color.FromArgb(125, 78, 0);
            _ruleNotice.BackColor = Color.FromArgb(255, 247, 220);
            _ruleNotice.AccessibleName = "规则状态：演示规则，不代表正式合规结论";
            _toolTip.SetToolTip(_ruleNotice,
                "当前 Word 插件连接的是 demo-0.1.0 演示规则，用于检查界面和修复流程；不是南开大学正式规则集，不能据此认定论文合规。");
            _ruleNotice.Paint += (sender, args) =>
            {
                using (var pen = new Pen(Color.FromArgb(242, 216, 155)))
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
            _issueList.IntegralHeight = false;
            _issueList.HorizontalScrollbar = false;
            _issueList.BorderStyle = BorderStyle.None;
            _issueList.BackColor = BackColor;
            _issueList.DrawMode = DrawMode.OwnerDrawVariable;
            _issueList.AccessibleName = "格式问题列表";
            _issueList.SelectedIndexChanged += OnIssueSelected;
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
                Padding = new Padding(9, 5, 9, 8)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 29F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));

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

            layout.Controls.Add(_summary, 0, 0);
            layout.Controls.Add(_ruleNotice, 0, 1);
            layout.Controls.Add(searchSurface, 0, 2);
            layout.Controls.Add(resultsSurface, 0, 3);
            layout.Controls.Add(_status, 0, 4);
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
        public bool CanResolveSelected => !_isBusy && SelectedDiagnostic != null;
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
            _analysisSucceeded = true;
            _diagnostics.Clear();
            if (result?.Diagnostics != null)
                _diagnostics.AddRange(result.Diagnostics.Where(diagnostic => diagnostic != null));
            _selectedIndex = -1;
            _selectedOccurrenceIndex = 0;
            ClearSearch();
            RebuildIssueList();
            string baseStatus = _diagnostics.Count == 0
                ? "未发现已覆盖范围内的格式问题；未覆盖项不会被判为合规。"
                : "悬停或点击正文黄色标记，可查看对应建议并定位原文。修复与验证操作位于顶部功能区。";
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

            int visibleOccurrences = _visibleIndices.Sum(index => GetOccurrenceCount(_diagnostics[index]));
            int totalOccurrences = _diagnostics.Sum(GetOccurrenceCount);
            _summary.Text = string.IsNullOrEmpty(query)
                ? "格式问题  " + _diagnostics.Count + " 项 · " + totalOccurrences + " 处"
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
                    case "期望": haystack = diagnostic.Expected; break;
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
                default:
                    if (!string.IsNullOrWhiteSpace(ruleId) && ruleId.IndexOf("caption", StringComparison.OrdinalIgnoreCase) >= 0)
                        return "图表题注";
                    return "其他格式";
            }
        }

        private void OnMeasureIssueItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _visibleIndices.Count)
            {
                e.ItemHeight = 100;
                return;
            }
            EngineDiagnostic diagnostic = _diagnostics[_visibleIndices[e.Index]];
            int width = Math.Max(160, _issueList.ClientSize.Width - 8);
            using (Graphics graphics = _issueList.CreateGraphics())
            {
                float textWidth = Math.Max(120, width - 86);
                int lines = 0;
                lines += MeasureWrappedLines(graphics, diagnostic.Expected, Font, textWidth);
                lines += MeasureWrappedLines(graphics, GetActualValues(diagnostic), Font, textWidth);
                lines += MeasureWrappedLines(graphics, GetLocations(diagnostic), Font, textWidth);
                lines += MeasureWrappedLines(graphics, BuildRecommendation(diagnostic), Font, textWidth);
                lines += MeasureWrappedLines(graphics, GetSource(diagnostic), Font, textWidth);
                e.ItemHeight = Math.Max(174, 64 + Math.Min(lines, 11) * 17 +
                    (GetOccurrenceCount(diagnostic) > 1 ? 15 : 0));
            }
        }

        private void OnDrawIssueItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index < 0 || e.Index >= _visibleIndices.Count) return;
            EngineDiagnostic diagnostic = _diagnostics[_visibleIndices[e.Index]];
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Rectangle card = Rectangle.Inflate(e.Bounds, -3, -3);
            card.Width = Math.Max(1, card.Width);
            card.Height = Math.Max(1, card.Height);
            Color fill = selected ? Color.FromArgb(255, 253, 225) : Color.White;
            Color border = selected ? Color.FromArgb(37, 112, 226) : Color.FromArgb(218, 225, 234);
            using (GraphicsPath path = RoundedRectangle(card, 7))
            using (var brush = new SolidBrush(fill))
            using (var pen = new Pen(border, selected ? 1.5F : 1F))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.FillPath(brush, path);
                e.Graphics.DrawPath(pen, path);
            }

            int left = card.Left + 12;
            int right = card.Right - 12;
            float y = card.Top + 10;
            using (var titleFont = new Font(Font, FontStyle.Bold))
            using (var smallFont = new Font(Font.FontFamily, 8.2F, FontStyle.Regular))
            using (var bodyFont = new Font(Font.FontFamily, 9F, FontStyle.Regular))
            using (var mutedBrush = new SolidBrush(Color.FromArgb(98, 111, 131)))
            using (var textBrush = new SolidBrush(ForeColor))
            using (var accentBrush = new SolidBrush(Color.FromArgb(39, 101, 191)))
            {
                string title = DescribeRule(diagnostic.RuleId);
                string count = GetOccurrenceCount(diagnostic) > 1 ? "共 " + GetOccurrenceCount(diagnostic) + " 处" : "1 处";
                e.Graphics.DrawString(title, titleFont, textBrush, left, y);
                SizeF countSize = e.Graphics.MeasureString(count, smallFont);
                e.Graphics.DrawString(count, smallFont, mutedBrush, Math.Max(left, right - countSize.Width), y + 2);
                y += 23;

                DrawCardLine(e.Graphics, "期望", diagnostic.Expected, left, ref y, right, bodyFont, accentBrush);
                DrawCardLine(e.Graphics, "实际", GetActualValues(diagnostic), left, ref y, right, bodyFont, textBrush);
                DrawCardLine(e.Graphics, "位置", GetLocations(diagnostic), left, ref y, right, smallFont, mutedBrush);
                DrawCardLine(e.Graphics, "建议", BuildRecommendation(diagnostic), left, ref y, right, smallFont, textBrush);
                DrawCardLine(e.Graphics, "依据", GetSource(diagnostic), left, ref y, right, smallFont, mutedBrush);
                y += 3;
            }
            e.DrawFocusRectangle();
        }

        private static void DrawCardLine(Graphics graphics, string label, string value, int left,
            ref float y, int right, Font font, Brush valueBrush)
        {
            using (var labelFont = new Font(font.FontFamily, font.Size, FontStyle.Bold))
            using (var labelBrush = new SolidBrush(Color.FromArgb(88, 101, 120)))
            {
                graphics.DrawString(label, labelFont, labelBrush, left, y);
                int labelWidth = (int)Math.Ceiling(graphics.MeasureString(label + "：", labelFont).Width) + 2;
                int textWidth = Math.Max(40, right - left - labelWidth);
                SizeF measured = graphics.MeasureString(value ?? "—", font, textWidth);
                Rectangle text = new Rectangle(left + labelWidth, (int)y, textWidth,
                    Math.Max(20, (int)Math.Ceiling(measured.Height) + 2));
                graphics.DrawString(value ?? "—", font, valueBrush, text);
                y += Math.Max(19, text.Height);
            }
        }

        private static int MeasureWrappedLines(Graphics graphics, string value, Font font, float width)
        {
            if (string.IsNullOrWhiteSpace(value)) return 1;
            SizeF size = graphics.MeasureString(value, font, (int)width);
            return Math.Max(1, (int)Math.Ceiling(size.Height / Math.Max(1F, font.GetHeight(graphics))));
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

        private static string GetSource(EngineDiagnostic diagnostic)
        {
            string message = diagnostic?.Message ?? string.Empty;
            int separator = message.IndexOf('：');
            if (separator > 0 && (diagnostic?.RuleId ?? string.Empty).StartsWith("NKU-REQ-", StringComparison.Ordinal))
                return message.Substring(0, separator);
            return "演示规则（仅用于验证插件流程，不代表正式标准）";
        }

        private static string BuildRecommendation(EngineDiagnostic diagnostic)
        {
            if (IsFixable(diagnostic))
                return "将所选格式改为「" + (diagnostic.Expected ?? "") + "」；修复前会先预览。";
            return "当前值无法安全自动修改，请依据规范核对后手动处理。";
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
