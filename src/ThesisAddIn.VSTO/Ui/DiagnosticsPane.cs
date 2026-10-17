using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>中文问题侧边栏：列出诊断、支持定位与修复操作。</summary>
    public sealed class DiagnosticsPane : UserControl
    {
        private readonly Button _analyzeButton = new Button();
        private readonly Button _fixOneButton = new Button();
        private readonly Button _fixAllButton = new Button();
        private readonly FlowLayoutPanel _diagnosticList = new FlowLayoutPanel();
        private readonly Panel _resultsHost = new Panel();
        private readonly Label _emptyState = new Label();
        private readonly TextBox _detail = new TextBox();
        private readonly Label _status = new Label();
        private readonly Label _coverage = new Label();
        private readonly List<EngineDiagnostic> _diagnostics = new List<EngineDiagnostic>();
        private readonly List<Button> _diagnosticCards = new List<Button>();
        private int _selectedIndex = -1;
        private bool _isBusy;

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
            _analyzeButton.Click += (s, e) => OnEvent(AnalyzeRequested);
            _fixOneButton.Click += (s, e) => OnEvent(FixCurrentRequested);
            _fixAllButton.Click += (s, e) => OnEvent(FixAllRequested);
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

            _resultsHost.Dock = DockStyle.Fill;
            _resultsHost.Controls.Add(_diagnosticList);
            _resultsHost.Controls.Add(_emptyState);
            _emptyState.BringToFront();

            _detail.Dock = DockStyle.Fill;
            _detail.Multiline = true;
            _detail.ReadOnly = true;
            _detail.ScrollBars = ScrollBars.Vertical;
            _detail.WordWrap = true;
            _detail.Text = "选择一条问题可定位到原文，并查看补充说明。";

            _status.Dock = DockStyle.Fill;
            _status.Padding = new Padding(2);
            _status.Text = "准备就绪。";

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(8)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            // 窄任务窗格下第三个操作按钮会换到第二行；保留足够高度，
            // 避免它被下方的覆盖范围说明遮住。
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 112F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 54F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 82F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));
            layout.Controls.Add(buttons, 0, 0);
            layout.Controls.Add(_coverage, 0, 1);
            layout.Controls.Add(_resultsHost, 0, 2);
            layout.Controls.Add(_detail, 0, 3);
            layout.Controls.Add(_status, 0, 4);
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
            _isBusy = busy;
            _analyzeButton.Enabled = !busy;
            _fixOneButton.Enabled = !busy && CanFixSelected;
            _fixAllButton.Enabled = !busy && HasAutoDiagnostics;
            if (!string.IsNullOrEmpty(message))
            {
                SetStatus(message);
            }
        }

        public void SetStatus(string message)
        {
            _status.Text = message;
        }

        public void ShowAnalysis(EngineAnalysisResult result)
        {
            _diagnostics.Clear();
            _diagnosticList.Controls.Clear();
            _diagnosticCards.Clear();
            _selectedIndex = -1;
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
            _diagnosticList.Visible = hasDiagnostics;
            _emptyState.Visible = !hasDiagnostics;
            if (!hasDiagnostics)
            {
                _emptyState.Text = "未发现问题。\n结果仅覆盖已检查范围。";
            }
            _detail.Text = "选择一条问题可定位到原文，并查看补充说明。";

            string ruleVersion = result != null ? result.RuleVersion : string.Empty;
            string coverage = result != null ? result.Coverage : string.Empty;
            _coverage.Text = "规则版本：" + (string.IsNullOrEmpty(ruleVersion) ? "（未提供）" : ruleVersion)
                + Environment.NewLine + "已检查范围：" + (string.IsNullOrEmpty(coverage) ? "（未提供）" : coverage);

            string uncovered = (result == null || result.UnsupportedScopes == null || result.UnsupportedScopes.Count == 0)
                ? "未覆盖范围：引擎未报告。本次仅覆盖演示范围内的少量规则，不代表论文合规。"
                : "未覆盖范围：" + string.Join("、", result.UnsupportedScopes.ToArray());
            SetStatus("完成分析，共 " + _diagnostics.Count + " 条问题。" + Environment.NewLine + uncovered);

            _fixOneButton.Enabled = CanFixSelected;
            _fixAllButton.Enabled = HasAutoDiagnostics;
        }

        public void ShowFailure(string message)
        {
            _diagnostics.Clear();
            _diagnosticList.Controls.Clear();
            _diagnosticCards.Clear();
            _diagnosticList.Visible = false;
            _selectedIndex = -1;
            _emptyState.Visible = true;
            _emptyState.Text = "本次分析未完成。";
            _fixOneButton.Enabled = false;
            _fixAllButton.Enabled = false;
            _coverage.Text = "本次分析未完成。";
            SetStatus(message);
        }

        /// <summary>
        /// 按批注中的规则、期望值和实际值选中唯一的问题。
        /// 多条诊断无法唯一对应时不激活卡片，避免跳到错误问题。
        /// </summary>
        public bool SelectByComment(string ruleId, string expected, string actual, bool navigateToSource)
        {
            if (string.IsNullOrWhiteSpace(ruleId))
            {
                return false;
            }

            int matchIndex = -1;
            for (int i = 0; i < _diagnostics.Count; i++)
            {
                EngineDiagnostic diagnostic = _diagnostics[i];
                if (string.Equals(diagnostic.RuleId, ruleId, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(diagnostic.Expected ?? string.Empty, expected ?? string.Empty, StringComparison.Ordinal)
                    && string.Equals(diagnostic.Actual ?? string.Empty, actual ?? string.Empty, StringComparison.Ordinal))
                {
                    if (matchIndex >= 0)
                    {
                        return false;
                    }

                    matchIndex = i;
                }
            }

            if (matchIndex < 0)
            {
                return false;
            }

            SelectDiagnostic(matchIndex, navigateToSource);
            return true;
        }

        private void AddDiagnosticCard(int index, EngineDiagnostic diagnostic)
        {
            var card = new Button
            {
                AutoSize = false,
                Height = 104,
                Text = BuildDiagnosticCardText(diagnostic),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.ControlText,
                AutoEllipsis = false,
                AccessibleRole = AccessibleRole.ListItem,
                AccessibleName = diagnostic.RuleId + "，" + diagnostic.Location
                    + "，期望值：" + diagnostic.Expected + "，实际值：" + diagnostic.Actual,
                Tag = index,
                Margin = new Padding(0, 0, 0, 6)
            };
            card.FlatAppearance.BorderSize = 1;
            card.FlatAppearance.BorderColor = SystemColors.ControlDark;
            card.Click += OnDiagnosticCardClick;
            _diagnosticCards.Add(card);
            _diagnosticList.Controls.Add(card);
            ResizeDiagnosticCards();
        }

        private static string BuildDiagnosticCardText(EngineDiagnostic diagnostic)
        {
            return diagnostic.RuleId + "  ·  " + DescribeSeverity(diagnostic.Severity)
                + " / " + DescribeFixLevel(diagnostic.FixLevel)
                + Environment.NewLine + "位置：" + (diagnostic.Location ?? "（未提供）")
                + Environment.NewLine + "期望：" + (diagnostic.Expected ?? "（未提供）")
                + Environment.NewLine + "实际：" + (diagnostic.Actual ?? "（未提供）");
        }

        private void ResizeDiagnosticCards()
        {
            int width = Math.Max(120, _diagnosticList.ClientSize.Width - _diagnosticList.Padding.Horizontal - 4);
            foreach (Button card in _diagnosticCards)
            {
                card.Width = width;
                int textWidth = Math.Max(80, width - card.Padding.Horizontal - 4);
                Size textSize = TextRenderer.MeasureText(
                    card.Text,
                    card.Font,
                    new Size(textWidth, int.MaxValue),
                    TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
                card.Height = Math.Max(104, textSize.Height + card.Padding.Vertical + 8);
            }
        }

        private void OnDiagnosticCardClick(object sender, EventArgs e)
        {
            Button card = sender as Button;
            if (card != null && card.Tag is int)
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
            for (int i = 0; i < _diagnosticCards.Count; i++)
            {
                Button card = _diagnosticCards[i];
                bool selected = i == index;
                card.BackColor = selected ? SystemColors.Info : SystemColors.Window;
                card.FlatAppearance.BorderColor = selected ? SystemColors.Highlight : SystemColors.ControlDark;
                card.AccessibleDescription = selected ? "当前选中" : string.Empty;
            }

            EngineDiagnostic diagnostic = SelectedDiagnostic;
            _detail.Text = string.IsNullOrWhiteSpace(diagnostic.Message)
                ? "该规则没有提供补充说明。"
                : diagnostic.Message;
            _fixOneButton.Enabled = !_isBusy && CanFixSelected;
            _diagnosticList.ScrollControlIntoView(_diagnosticCards[index]);
            if (navigateToSource)
            {
                OnEvent(DiagnosticSelected);
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

