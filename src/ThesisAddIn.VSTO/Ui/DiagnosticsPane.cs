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
        private readonly ListView _list = new ListView();
        private readonly TextBox _detail = new TextBox();
        private readonly Label _status = new Label();
        private readonly Label _coverage = new Label();
        private readonly List<EngineDiagnostic> _diagnostics = new List<EngineDiagnostic>();

        public event EventHandler AnalyzeRequested;
        public event EventHandler FixCurrentRequested;
        public event EventHandler FixAllRequested;
        public event EventHandler DiagnosticSelected;

        public DiagnosticsPane()
        {
            Dock = DockStyle.Fill;
            Padding = new Padding(8);
            Font = new Font("Microsoft YaHei UI", 9F);

            // Word 的任务窗格很窄，三个按钮必须允许换行，否则最后一个会被裁掉。
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 68,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                AutoScroll = false
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

            _coverage.Dock = DockStyle.Top;
            _coverage.Height = 34;
            _coverage.Text = "尚未分析。";

            _list.Dock = DockStyle.Fill;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.MultiSelect = false;
            _list.HideSelection = false;
            _list.Columns.Add("序号", 40);
            _list.Columns.Add("规则", 150);
            _list.Columns.Add("位置", 130);
            _list.Columns.Add("期望值", 110);
            _list.Columns.Add("实际值", 110);
            _list.Columns.Add("严重度", 60);
            _list.Columns.Add("处理", 60);
            _list.SelectedIndexChanged += OnListSelectionChanged;

            _detail.Dock = DockStyle.Bottom;
            _detail.Height = 96;
            _detail.Multiline = true;
            _detail.ReadOnly = true;
            _detail.ScrollBars = ScrollBars.Vertical;
            _detail.Text = "选择上方的一条问题，可以查看详情并定位到原文。";

            _status.Dock = DockStyle.Bottom;
            _status.Height = 52;
            _status.Text = "准备就绪。";

            Controls.Add(_list);
            Controls.Add(_detail);
            Controls.Add(_status);
            Controls.Add(_coverage);
            Controls.Add(buttons);
        }

        public EngineDiagnostic SelectedDiagnostic
        {
            get
            {
                int index = _list.SelectedIndices.Count > 0 ? _list.SelectedIndices[0] : -1;
                return index >= 0 && index < _diagnostics.Count ? _diagnostics[index] : null;
            }
        }

        public IList<EngineDiagnostic> Diagnostics
        {
            get { return _diagnostics; }
        }

        public void SetBusy(bool busy, string message)
        {
            _analyzeButton.Enabled = !busy;
            _fixOneButton.Enabled = !busy && HasAutoDiagnostics;
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
            _list.Items.Clear();
            if (result != null && result.Diagnostics != null)
            {
                foreach (EngineDiagnostic diagnostic in result.Diagnostics)
                {
                    _diagnostics.Add(diagnostic);
                    var item = new ListViewItem((_diagnostics.Count).ToString());
                    item.SubItems.Add(diagnostic.RuleId);
                    item.SubItems.Add(diagnostic.Location);
                    item.SubItems.Add(diagnostic.Expected);
                    item.SubItems.Add(diagnostic.Actual);
                    item.SubItems.Add(DescribeSeverity(diagnostic.Severity));
                    item.SubItems.Add(DescribeFixLevel(diagnostic.FixLevel));
                    _list.Items.Add(item);
                }
            }

            string ruleVersion = result != null ? result.RuleVersion : string.Empty;
            string coverage = result != null ? result.Coverage : string.Empty;
            _coverage.Text = "规则版本：" + (string.IsNullOrEmpty(ruleVersion) ? "（未提供）" : ruleVersion)
                + Environment.NewLine + "已检查范围：" + (string.IsNullOrEmpty(coverage) ? "（未提供）" : coverage);

            string uncovered = (result == null || result.UnsupportedScopes == null || result.UnsupportedScopes.Count == 0)
                ? "未覆盖范围：引擎未报告。本次仅覆盖演示范围内的少量规则，不代表论文合规。"
                : "未覆盖范围：" + string.Join("、", result.UnsupportedScopes.ToArray());
            SetStatus("完成分析，共 " + _diagnostics.Count + " 条问题。" + Environment.NewLine + uncovered);

            _fixOneButton.Enabled = HasAutoDiagnostics;
            _fixAllButton.Enabled = HasAutoDiagnostics;
        }

        public void ShowFailure(string message)
        {
            _diagnostics.Clear();
            _list.Items.Clear();
            _fixOneButton.Enabled = false;
            _fixAllButton.Enabled = false;
            _coverage.Text = "本次分析未完成。";
            SetStatus(message);
        }

        /// <summary>按规则编号选中对应的问题（用于批注锚点联动）。</summary>
        public bool SelectByRuleId(string ruleId)
        {
            if (string.IsNullOrWhiteSpace(ruleId))
            {
                return false;
            }

            for (int i = 0; i < _diagnostics.Count; i++)
            {
                if (string.Equals(_diagnostics[i].RuleId, ruleId, StringComparison.OrdinalIgnoreCase))
                {
                    _list.Items[i].Selected = true;
                    _list.Items[i].EnsureVisible();
                    _list.Focus();
                    return true;
                }
            }

            return false;
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

        private void OnListSelectionChanged(object sender, EventArgs e)
        {
            EngineDiagnostic diagnostic = SelectedDiagnostic;
            if (diagnostic == null)
            {
                return;
            }

            _detail.Text = string.Join(Environment.NewLine, new[]
            {
                "规则：" + diagnostic.RuleId + "（严重度：" + DescribeSeverity(diagnostic.Severity) + "，处理方式：" + DescribeFixLevel(diagnostic.FixLevel) + "）",
                "位置：" + diagnostic.Location,
                "期望值：" + diagnostic.Expected,
                "实际值：" + diagnostic.Actual,
                "说明：" + diagnostic.Message
            });
            OnEvent(DiagnosticSelected);
        }

        private static void ConfigureButton(Button button, string text)
        {
            button.Text = text;
            button.AutoSize = true;
            button.Margin = new Padding(0, 0, 6, 0);
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
