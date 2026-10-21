using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>紧凑的中文任务窗格：联动原文批注、定位诊断并提供修复操作。</summary>
    public sealed class DiagnosticsPane : UserControl
    {
        private readonly Button _analyzeButton = new Button();
        private readonly Button _fixOneButton = new Button();
        private readonly Button _fixAllButton = new Button();
        private readonly Button _previousButton = new Button();
        private readonly Button _nextButton = new Button();
        private readonly Button _toggleListButton = new Button();
        private readonly FlowLayoutPanel _diagnosticList = new FlowLayoutPanel();
        private readonly Panel _resultsHost = new Panel();
        private readonly Label _emptyState = new Label();
        private readonly Label _issueSummary = new Label();
        private readonly Label _status = new Label();
        private readonly Label _coverage = new Label();
        private readonly List<EngineDiagnostic> _diagnostics = new List<EngineDiagnostic>();
        private readonly List<Panel> _diagnosticCards = new List<Panel>();
        private int _selectedIndex = -1;
        private bool _isBusy;
        private bool _showDiagnosticList;

        // 当前演示规则集依据南开大学 2026 版规范第 4.7 节表 4.6。
        private const string DemoRequirementSource = "2026版第4.7节表4.6";

        public event EventHandler AnalyzeRequested;
        public event EventHandler FixCurrentRequested;
        public event EventHandler FixAllRequested;
        public event EventHandler DiagnosticSelected;

        public DiagnosticsPane()
        {
            Dock = DockStyle.Fill;
            Font = new Font("Microsoft YaHei UI", 9F);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                AutoScroll = true,
                Padding = new Padding(0, 2, 0, 0)
            };
            ConfigureButton(_analyzeButton, "分析当前文档");
            ConfigureButton(_fixOneButton, "修复此项");
            ConfigureButton(_fixAllButton, "预览并修复全部");
            _fixOneButton.Enabled = false;
            _fixAllButton.Enabled = false;
            _previousButton.Enabled = false;
            _nextButton.Enabled = false;
            _toggleListButton.Enabled = false;
            _analyzeButton.Click += (s, e) => OnEvent(AnalyzeRequested);
            _fixOneButton.Click += (s, e) => OnEvent(FixCurrentRequested);
            _fixAllButton.Click += (s, e) => OnEvent(FixAllRequested);
            _previousButton.Click += (s, e) => NavigateSelection(-1);
            _nextButton.Click += (s, e) => NavigateSelection(1);
            _toggleListButton.Click += (s, e) => ToggleDiagnosticList();
            buttons.Controls.Add(_analyzeButton);
            buttons.Controls.Add(_fixOneButton);
            buttons.Controls.Add(_fixAllButton);

            _coverage.Dock = DockStyle.Fill;
            _coverage.Padding = new Padding(2, 2, 2, 2);
            _coverage.AutoEllipsis = true;
            _coverage.Text = "尚未分析。";

            // Word 任务窗格通常只有较窄的宽度。卡片将关键信息分行显示，
            // 避免固定多列表格把期望值和实际值挤到横向滚动区域之外。
            _diagnosticList.Dock = DockStyle.Fill;
            _diagnosticList.FlowDirection = FlowDirection.TopDown;
            _diagnosticList.WrapContents = false;
            _diagnosticList.AutoScroll = true;
            _diagnosticList.Padding = new Padding(2);
            _diagnosticList.SizeChanged += (s, e) => ResizeDiagnosticCards();

            _emptyState.Dock = DockStyle.Fill;
            _emptyState.TextAlign = ContentAlignment.MiddleCenter;
            _emptyState.ForeColor = SystemColors.GrayText;
            _emptyState.Padding = new Padding(12);
            _emptyState.Text = "点击“分析当前文档”开始检查。";

            _issueSummary.Dock = DockStyle.Fill;
            _issueSummary.Padding = new Padding(2, 6, 2, 2);
            _issueSummary.AutoEllipsis = true;
            _issueSummary.Text = "分析后可展开问题列表；已有工具批注与原文保持联动。";

            var navigation = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                AutoScroll = true,
                Padding = new Padding(0, 2, 0, 0)
            };
            ConfigureButton(_previousButton, "上一项");
            ConfigureButton(_nextButton, "下一项");
            ConfigureButton(_toggleListButton, "展开问题列表");
            navigation.Controls.Add(_previousButton);
            navigation.Controls.Add(_nextButton);
            navigation.Controls.Add(_toggleListButton);

            _resultsHost.Dock = DockStyle.Fill;
            _resultsHost.Controls.Add(_diagnosticList);
            _resultsHost.Controls.Add(_emptyState);
            _emptyState.BringToFront();

            _status.Dock = DockStyle.Fill;
            _status.Padding = new Padding(2);
            _status.Text = "准备就绪。";

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(8)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            // 操作按钮紧凑置顶；问题详情优先由原文旁的 Word 批注呈现。
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 92F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
            layout.Controls.Add(buttons, 0, 0);
            layout.Controls.Add(_coverage, 0, 1);
            layout.Controls.Add(navigation, 0, 2);
            layout.Controls.Add(_issueSummary, 0, 3);
            layout.Controls.Add(_resultsHost, 0, 4);
            layout.Controls.Add(_status, 0, 5);
            Controls.Add(layout);
        }

        public EngineDiagnostic SelectedDiagnostic
        {
            get
            {
                return _selectedIndex >= 0 && _selectedIndex < _diagnostics.Count
                    ? _diagnostics[_selectedIndex]
                    : null;
            }
        }

        public IList<EngineDiagnostic> Diagnostics
        {
            get { return _diagnostics; }
        }

        public void SetBusy(bool busy, string message)
        {
            InvokeOnUiThread(() =>
            {
                _isBusy = busy;
                _analyzeButton.Enabled = !busy;
                _fixOneButton.Enabled = !busy && CanFixSelected;
                _fixAllButton.Enabled = !busy && HasAutoDiagnostics;
                _previousButton.Enabled = !busy && _diagnostics.Count > 0;
                _nextButton.Enabled = !busy && _diagnostics.Count > 0;
                _toggleListButton.Enabled = !busy && _diagnostics.Count > 0;
                if (!string.IsNullOrEmpty(message))
                {
                    _status.Text = message;
                }
            });
        }

        public void SetStatus(string message)
        {
            InvokeOnUiThread(() =>
            {
                _status.Text = message;
            });
        }

        public void ShowAnalysis(EngineAnalysisResult result)
        {
            InvokeOnUiThread(() => ShowAnalysisCore(result));
        }

        private void ShowAnalysisCore(EngineAnalysisResult result)
        {
            _diagnostics.Clear();
            _diagnosticList.Controls.Clear();
            _diagnosticCards.Clear();
            _selectedIndex = -1;
            _showDiagnosticList = false;
            if (result != null && result.Diagnostics != null)
            {
                foreach (EngineDiagnostic diagnostic in result.Diagnostics)
                {
                    _diagnostics.Add(diagnostic);
                }
            }

            for (int i = 0; i < _diagnostics.Count; i++)
            {
                AddDiagnosticCard(i, _diagnostics[i]);
            }
            bool hasDiagnostics = _diagnostics.Count > 0;
            _diagnosticList.Visible = hasDiagnostics && _showDiagnosticList;
            _emptyState.Visible = !hasDiagnostics;
            if (!hasDiagnostics)
            {
                _emptyState.Text = "未发现问题。\n结果仅覆盖已检查范围。";
            }
            _previousButton.Enabled = hasDiagnostics;
            _nextButton.Enabled = hasDiagnostics;
            _toggleListButton.Enabled = hasDiagnostics;
            _toggleListButton.Text = "展开问题列表";
            _issueSummary.Text = hasDiagnostics
                ? "检测到 " + _diagnostics.Count + " 条问题。可用上一项/下一项浏览，或展开问题列表查看详情。"
                : "没有待处理问题。结果仅覆盖已检查范围。";
            string ruleVersion = result != null ? result.RuleVersion : string.Empty;
            string coverage = result != null ? result.Coverage : string.Empty;
            _coverage.Text = "规则版本：" + (string.IsNullOrEmpty(ruleVersion) ? "（未提供）" : ruleVersion)
                + "；已检查：" + (string.IsNullOrEmpty(coverage) ? "（未提供）" : coverage);

            string uncovered = (result == null || result.UnsupportedScopes == null || result.UnsupportedScopes.Count == 0)
                ? "未覆盖：引擎未报告；本演示只覆盖少量规则，不代表论文合规。"
                : "未覆盖范围：" + string.Join("、", result.UnsupportedScopes.ToArray());
            _coverage.Text += Environment.NewLine + uncovered;
            SetStatus("分析完成，共 " + _diagnostics.Count + " 条问题。悬停正文高亮可联动建议。");

            _fixOneButton.Enabled = CanFixSelected;
            _fixAllButton.Enabled = HasAutoDiagnostics;
        }

        public void ShowFailure(string message)
        {
            InvokeOnUiThread(() => ShowFailureCore(message));
        }

        private void ShowFailureCore(string message)
        {
            _diagnostics.Clear();
            _diagnosticList.Controls.Clear();
            _diagnosticCards.Clear();
            _diagnosticList.Visible = false;
            _selectedIndex = -1;
            _showDiagnosticList = false;
            _emptyState.Visible = true;
            string detail = string.IsNullOrWhiteSpace(message) ? "未知错误。" : message.Trim();
            _emptyState.Text = "本次分析未完成。" + Environment.NewLine + Environment.NewLine + detail;
            _fixOneButton.Enabled = false;
            _fixAllButton.Enabled = false;
            _previousButton.Enabled = false;
            _nextButton.Enabled = false;
            _toggleListButton.Enabled = false;
            _toggleListButton.Text = "展开问题列表";
            _issueSummary.Text = "分析失败；请查看下方错误说明。";
            _coverage.Text = "本次分析未完成。";
            _status.Text = "本次分析未完成，错误详情见结果区域。";
        }

        /// <summary>
        /// Engine protocol calls use ConfigureAwait(false), so their callers may resume
        /// on a worker thread. Keep all WinForms control access on the pane's UI thread.
        /// Invoke is synchronous so the async operation observes UI exceptions and keeps
        /// updates ordered with the following status changes.
        /// </summary>
        public T InvokeOnUiThread<T>(Func<T> action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            if (IsDisposed || Disposing || !IsHandleCreated)
            {
                throw new InvalidOperationException("论文格式检查面板已关闭，无法更新界面。");
            }

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

        /// <summary>
        /// 按批注中的规则、期望值和实际值选中唯一的问题。
        /// 多条诊断无法唯一对应时不激活卡片，避免跳到错误问题。
        /// </summary>
        public bool SelectByComment(string ruleId, string expected, string actual, string location, bool navigateToSource)
        {
            if (string.IsNullOrWhiteSpace(ruleId))
            {
                return false;
            }

            int matchIndex = CommentDiagnosticMatcher.FindUniqueMatchIndex(
                _diagnostics, ruleId, expected, actual, location);
            if (matchIndex < 0)
            {
                return false;
            }

            SelectDiagnostic(matchIndex, navigateToSource);
            return true;
        }

        private void AddDiagnosticCard(int index, EngineDiagnostic diagnostic)
        {
            var card = new Panel
            {
                Height = 120,
                Padding = new Padding(0),
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                AccessibleRole = AccessibleRole.ListItem,
                AccessibleName = BuildDiagnosticCardText(diagnostic),
                Tag = index,
                Margin = new Padding(0, 0, 0, 8),
                Cursor = Cursors.Hand
            };

            var avatar = new Panel
            {
                Name = "Avatar",
                Left = 10,
                Top = 12,
                Width = 36,
                Height = 36,
                BackColor = SystemColors.Window,
                Cursor = Cursors.Hand
            };
            avatar.Paint += (sender, args) =>
            {
                args.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var brush = new SolidBrush(Color.FromArgb(211, 50, 57)))
                using (var font = new Font("Microsoft YaHei UI", 8F, FontStyle.Bold))
                using (var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                })
                {
                    args.Graphics.FillEllipse(brush, 1, 1, avatar.Width - 2, avatar.Height - 2);
                    args.Graphics.DrawString("NT", font, Brushes.White,
                        new RectangleF(1, 1, avatar.Width - 2, avatar.Height - 2), format);
                }
            };

            var author = new Label
            {
                Name = "Author",
                Text = "NKU Thesis Tool",
                Left = 56,
                Top = 10,
                Height = 22,
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = SystemColors.ControlText,
                BackColor = Color.Transparent,
                AutoEllipsis = true,
                Cursor = Cursors.Hand
            };

            var body = new Label
            {
                Name = "Body",
                Text = BuildDiagnosticCardText(diagnostic),
                Left = 56,
                Top = 34,
                AutoSize = false,
                Font = new Font("Microsoft YaHei UI", 9F),
                ForeColor = SystemColors.ControlText,
                BackColor = Color.Transparent,
                UseMnemonic = false,
                Cursor = Cursors.Hand
            };

            card.Controls.Add(avatar);
            card.Controls.Add(author);
            card.Controls.Add(body);
            AttachCardClickHandlers(card);
            _diagnosticCards.Add(card);
            _diagnosticList.Controls.Add(card);
            ResizeDiagnosticCards();
        }

        private static string BuildDiagnosticCardText(EngineDiagnostic diagnostic)
        {
            return "要求参考来源：" + DemoRequirementSource
                + Environment.NewLine + "规则：" + (diagnostic.RuleId ?? "（未提供）")
                + Environment.NewLine + "期望值：" + (diagnostic.Expected ?? "（未提供）")
                + Environment.NewLine + "实际值：" + (diagnostic.Actual ?? "（未提供）");
        }

        private void ResizeDiagnosticCards()
        {
            int width = Math.Max(120, _diagnosticList.ClientSize.Width - _diagnosticList.Padding.Horizontal - 4);
            foreach (Panel card in _diagnosticCards)
            {
                card.Width = width;
                Label author = card.Controls["Author"] as Label;
                Label body = card.Controls["Body"] as Label;
                if (author == null || body == null)
                {
                    continue;
                }

                int textWidth = Math.Max(80, width - body.Left - 12);
                author.Width = textWidth;
                body.Width = textWidth;
                Size textSize = TextRenderer.MeasureText(
                    body.Text,
                    body.Font,
                    new Size(textWidth, int.MaxValue),
                    TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
                // TextRenderer can measure wrapped Chinese text one line shorter than
                // the WinForms Label paints it at the pane's DPI. Reserve one extra
                // line so the final “实际值” row is never clipped by the card border.
                body.Height = textSize.Height + body.Font.Height + 2;
                card.Height = Math.Max(82, body.Top + body.Height + 12);
            }
        }

        private void AttachCardClickHandlers(Control control)
        {
            control.Click += OnDiagnosticCardClick;
            foreach (Control child in control.Controls)
            {
                AttachCardClickHandlers(child);
            }
        }

        private void OnDiagnosticCardClick(object sender, EventArgs e)
        {
            Control control = sender as Control;
            while (control != null && !(control.Tag is int))
            {
                control = control.Parent;
            }

            if (control is Panel card && card.Tag is int)
            {
                SelectDiagnostic((int)card.Tag);
            }
        }

        private void SelectDiagnostic(int index)
        {
            SelectDiagnostic(index, navigateToSource: true);
        }

        private void SelectDiagnostic(int index, bool navigateToSource)
        {
            if (index < 0 || index >= _diagnostics.Count)
            {
                return;
            }

            _selectedIndex = index;
            EngineDiagnostic selectedDiagnostic = _diagnostics[index];
            _issueSummary.Text = "当前 " + (index + 1) + "/" + _diagnostics.Count + "："
                + (selectedDiagnostic.RuleId ?? "（未提供规则）") + "。已有批注可在原文查看，或展开问题列表查看详情。";
            for (int i = 0; i < _diagnosticCards.Count; i++)
            {
                Panel card = _diagnosticCards[i];
                bool selected = i == index;
                card.BackColor = selected ? SystemColors.Info : SystemColors.Window;
                card.BorderStyle = BorderStyle.FixedSingle;
                card.AccessibleDescription = selected ? "当前选中" : string.Empty;
            }

            _fixOneButton.Enabled = !_isBusy && CanFixSelected;
            if (_showDiagnosticList && index < _diagnosticCards.Count)
            {
                _diagnosticList.ScrollControlIntoView(_diagnosticCards[index]);
            }
            if (navigateToSource)
            {
                OnEvent(DiagnosticSelected);
            }
        }

        private void NavigateSelection(int direction)
        {
            if (_isBusy || _diagnostics.Count == 0 || direction == 0)
            {
                return;
            }

            int nextIndex;
            if (_selectedIndex < 0)
            {
                nextIndex = direction > 0 ? 0 : _diagnostics.Count - 1;
            }
            else
            {
                nextIndex = (_selectedIndex + direction + _diagnostics.Count) % _diagnostics.Count;
            }

            SelectDiagnostic(nextIndex);
        }

        private void ToggleDiagnosticList()
        {
            if (_isBusy || _diagnostics.Count == 0)
            {
                return;
            }

            _showDiagnosticList = !_showDiagnosticList;
            _diagnosticList.Visible = _showDiagnosticList;
            _emptyState.Visible = false;
            _toggleListButton.Text = _showDiagnosticList ? "收起问题列表" : "展开问题列表";
            if (_showDiagnosticList && _selectedIndex >= 0 && _selectedIndex < _diagnosticCards.Count)
            {
                _diagnosticList.ScrollControlIntoView(_diagnosticCards[_selectedIndex]);
            }
        }

        private bool HasAutoDiagnostics
        {
            get
            {
                foreach (EngineDiagnostic diagnostic in _diagnostics)
                {
                    if (string.Equals(diagnostic.FixLevel, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        private bool CanFixSelected
        {
            get
            {
                EngineDiagnostic diagnostic = SelectedDiagnostic;
                return diagnostic != null
                    && string.Equals(diagnostic.FixLevel, "auto", StringComparison.OrdinalIgnoreCase);
            }
        }

        private static void ConfigureButton(Button button, string text)
        {
            button.Text = text;
            button.AutoSize = true;
            button.Height = 30;
            button.Margin = new Padding(0, 0, 6, 4);
        }

        private void OnEvent(EventHandler handler)
        {
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

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



