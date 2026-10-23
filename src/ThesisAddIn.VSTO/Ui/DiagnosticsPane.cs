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
    /// <summary>只显示可检索的问题列表；分析、定位和修复命令位于 Word 功能区。</summary>
    public sealed class DiagnosticsPane : UserControl
    {
        private const string SearchPlaceholder = "搜索问题（例：类型:西文字体 期望:Times New Roman）";
        private readonly TextBox _searchBox = new TextBox();
        private readonly Button _clearSearchButton = new Button();
        private readonly StableIssueListBox _issueList = new StableIssueListBox();
        private readonly Label _emptyResults = new Label();
        private readonly Label _summary = new Label();
        private readonly Label _ruleNotice = new Label();
        private readonly Label _status = new Label();
        private readonly ToolTip _toolTip = new ToolTip();
        private readonly Timer _navigationTimer = new Timer { Interval = 100 };
        private readonly Image _headerImage;
        private readonly List<EngineDiagnostic> _diagnostics = new List<EngineDiagnostic>();
        private readonly List<int> _visibleIndices = new List<int>();
        private int _selectedIndex = -1;
        private int _selectedOccurrenceIndex;
        private bool _isBusy;
        private bool _isShowingPlaceholder;
        private bool _updatingList;
        private bool _analysisSucceeded;
        private float _uiScale = 1F;
        private readonly List<Font> _scaledFonts = new List<Font>();

        public event EventHandler DiagnosticSelected;
        public event EventHandler StateChanged;

        public DiagnosticsPane()
        {
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
                Padding = new Padding(9, 6, 9, 8),
                BackColor = BackColor
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 86F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 31F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));

            _headerImage = LoadHeaderImage();
            var header = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(248, 242, 233),
                Margin = new Padding(0, 0, 0, 4),
                Padding = new Padding(0)
            };
            header.Paint += (sender, args) => PaintHeader(args.Graphics, header.ClientRectangle);
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
            layout.Controls.Add(_ruleNotice, 0, 1);
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
            return (int)Math.Round(390F * _uiScale);
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
                if (busy) _navigationTimer.Stop();
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
            _navigationTimer.Stop();
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
                    int firstVisible = _issueList.TopIndex;
                    if (_issueList.SelectedIndex != visibleIndex)
                        _issueList.SelectedIndex = visibleIndex;
                    // Body/comment hover must not scroll the searchable index either.
                    if (firstVisible >= 0 && firstVisible < _issueList.Items.Count &&
                        _issueList.TopIndex != firstVisible) _issueList.TopIndex = firstVisible;
                }
                finally { _updatingList = false; }
            }
            if (priorIndex != _selectedIndex || priorOccurrence != _selectedOccurrenceIndex)
            {
                if (visibleIndex >= 0) _issueList.Invalidate(_issueList.GetItemRectangle(visibleIndex));
                RaiseStateChanged();
            }
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
                case "demo/caption-east-asian-font": return "图表题注字体";
                case "demo/caption-font-size": return "图表题注字号";
                case "demo/caption-alignment": return "图表题注对齐";
                case "demo/caption-line-spacing": return "图表题注行距";
                case "demo/caption-space-before": return "图表题注段前距";
                case "demo/caption-space-after": return "图表题注段后距";
                case "demo/caption-placement": return "图表题注位置";
                case "demo/caption-number-gap": return "题注编号间距";
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
                e.ItemHeight = (int)Math.Ceiling(100 * _uiScale);
                return;
            }
            EngineDiagnostic diagnostic = _diagnostics[_visibleIndices[e.Index]];
            int width = Math.Max(160, (int)(_issueList.ClientSize.Width / _uiScale) - 8);
            using (Graphics graphics = _issueList.CreateGraphics())
            {
                int left = 15;
                int right = width - 15;
                int height = 14 + 24 + 4;
                using (var smallFont = LogicalFont(8.4F))
                using (var bodyFont = LogicalFont(9F))
                {
                    height += MeasureCardLineHeight(graphics, diagnostic.Expected, left, right, bodyFont) + 4;
                    height += MeasureCardLineHeight(graphics, GetActualValues(diagnostic), left, right, bodyFont) + 4;
                    height += MeasureCardLineHeight(graphics,
                        GetCompactLocation(_visibleIndices[e.Index], diagnostic), left, right, smallFont) + 4;
                    height += MeasureCardLineHeight(graphics, BuildRecommendation(diagnostic), left, right, smallFont) + 10;
                }
                e.ItemHeight = (int)Math.Ceiling(Math.Max(126, height) * _uiScale);
            }
        }

        private void OnDrawIssueItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index < 0 || e.Index >= _visibleIndices.Count) return;
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            GraphicsState state = e.Graphics.Save();
            try
            {
                e.Graphics.ScaleTransform(_uiScale, _uiScale);
                var bounds = new Rectangle((int)(e.Bounds.X / _uiScale), (int)(e.Bounds.Y / _uiScale),
                    (int)(e.Bounds.Width / _uiScale), (int)(e.Bounds.Height / _uiScale));
                DrawIssueCard(e.Graphics, bounds, e.Index, selected);
            }
            finally { e.Graphics.Restore(state); }
            e.DrawFocusRectangle();
        }

        private void DrawIssueCard(Graphics graphics, Rectangle bounds, int index, bool selected)
        {
            EngineDiagnostic diagnostic = _diagnostics[_visibleIndices[index]];
            Rectangle card = Rectangle.Inflate(bounds, -3, -3);
            card.Width = Math.Max(1, card.Width);
            card.Height = Math.Max(1, card.Height);
            Color fill = selected ? Color.FromArgb(255, 249, 233) : Color.White;
            Color border = selected ? Color.FromArgb(132, 38, 76) : Color.FromArgb(225, 218, 211);
            using (GraphicsPath path = RoundedRectangle(card, 7))
            using (var brush = new SolidBrush(fill))
            using (var pen = new Pen(border, selected ? 1.5F : 1F))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.FillPath(brush, path);
                graphics.DrawPath(pen, path);
            }
            if (selected)
            {
                using (var accent = new SolidBrush(Color.FromArgb(132, 38, 76)))
                    graphics.FillRectangle(accent, card.Left + 1, card.Top + 8, 3, Math.Max(12, card.Height - 16));
            }

            int left = card.Left + 12;
            int right = card.Right - 12;
            float y = card.Top + 10;
            using (var titleFont = LogicalFont(9.5F, FontStyle.Bold))
            using (var smallFont = LogicalFont(8.4F))
            using (var bodyFont = LogicalFont(9F))
            using (var mutedBrush = new SolidBrush(Color.FromArgb(112, 99, 102)))
            using (var textBrush = new SolidBrush(ForeColor))
            using (var accentBrush = new SolidBrush(Color.FromArgb(111, 34, 66)))
            {
                string title = DescribeRule(diagnostic.RuleId);
                string count = GetOccurrenceCount(diagnostic) > 1 ? "共 " + GetOccurrenceCount(diagnostic) + " 处" : "1 处";
                using (var titleBrush = new SolidBrush(Color.FromArgb(104, 31, 61)))
                    graphics.DrawString(title, titleFont, titleBrush, left, y);
                SizeF countSize = graphics.MeasureString(count, smallFont);
                graphics.DrawString(count, smallFont, mutedBrush, Math.Max(left, right - countSize.Width), y + 2);
                y += 25;

                DrawCardLine(graphics, "期望", diagnostic.Expected, left, ref y, right, bodyFont, accentBrush);
                DrawCardLine(graphics, "实际", GetActualValues(diagnostic), left, ref y, right, bodyFont, textBrush);
                DrawCardLine(graphics, "位置", GetCompactLocation(_visibleIndices[index], diagnostic),
                    left, ref y, right, smallFont, mutedBrush);
                DrawCardLine(graphics, "建议", BuildRecommendation(diagnostic), left, ref y, right, smallFont, textBrush);
            }
        }

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
                y += rowHeight + 4;
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
            return "当前第 " + (occurrenceIndex + 1) + "/" + count + " 处：" + description
                + "；可用功能区“上一处/下一处”浏览。";
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
