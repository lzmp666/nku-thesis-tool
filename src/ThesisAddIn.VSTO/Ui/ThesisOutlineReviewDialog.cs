using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>
    /// Reviews structure candidates without changing Word or accepting model suggestions implicitly.
    /// The caller must validate the returned approvals against the original document fingerprint.
    /// </summary>
    internal sealed class ThesisOutlineReviewDialog : Form
    {
        private const string ConfirmColumn = "confirmed";
        private const string SectionColumn = "section";
        private const string KindColumn = "kind";
        private const string LevelColumn = "level";
        private const string StateColumn = "state";
        private static readonly Color Purple = Color.FromArgb(104, 31, 61);
        private readonly DataGridView _grid = new DataGridView();
        private readonly ComboBox _sectionFilter = new ComboBox();
        private readonly ComboBox _titleMode = new ComboBox();
        private readonly TextBox _search = new TextBox();
        private readonly TextBox _detail = new TextBox();
        private readonly Label _summary = new Label();
        private readonly Label _notice = new Label();
        private readonly CheckBox _acknowledgment = new CheckBox();
        private readonly Button _confirm = new Button();
        private readonly Button _confirmVisible = new Button();
        private readonly List<RowState> _rows = new List<RowState>();
        private readonly List<Font> _ownedFonts = new List<Font>();
        private readonly Timer _filterTimer = new Timer { Interval = 160 };
        private bool _updating;
        private bool _scaledForOffice;

        internal event Action<EngineStructureBlock> LocateRequested;
        internal List<EngineApprovedStructureBlock> ApprovedBlocks { get; private set; }
            = new List<EngineApprovedStructureBlock>();
        internal IReadOnlyList<EngineApprovedStructureBlock> ConfirmedRoles => ApprovedBlocks.AsReadOnly();
        internal EngineThesisTitleMode TitleMode => _titleMode.SelectedIndex == 1
            ? EngineThesisTitleMode.Mode2 : EngineThesisTitleMode.Mode1;

        internal ThesisOutlineReviewDialog(EngineStructureReviewResult review)
        {
            if (review == null) throw new ArgumentNullException(nameof(review));
            if (review.Blocks == null) throw new ArgumentException("大纲块列表缺失。", nameof(review));
            if (review.Blocks.Any(block => block == null || string.IsNullOrWhiteSpace(block.BlockId)) ||
                review.Blocks.Select(block => block.BlockId).Distinct(StringComparer.Ordinal).Count() != review.Blocks.Count)
                throw new ArgumentException("大纲存在空或重复的原文锚点。", nameof(review));

            Text = "南开 · 整篇大纲确认";
            AutoScaleMode = AutoScaleMode.None;
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(1100, 770);
            MinimumSize = new Size(900, 620);
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.FromArgb(250, 247, 241);
            ForeColor = Color.FromArgb(44, 39, 42);
            Font = OwnFont(9F);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6,
                Padding = new Padding(18, 12, 18, 12), BackColor = BackColor
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 67));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 78));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 39));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 108));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 91));

            var heading = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute, 31));
            heading.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            heading.Controls.Add(new Label
            {
                Text = "先核对大纲，再按角色排版", Dock = DockStyle.Fill,
                Font = OwnFont(16F, FontStyle.Bold), ForeColor = Purple
            }, 0, 0);
            heading.Controls.Add(new Label
            {
                Text = "每行对应原文一个段落。可修改所属部分、类型和标题层级；点击“定位”核对原文。所有候选初始均未确认。",
                Dock = DockStyle.Fill, Padding = new Padding(0, 3, 0, 0)
            }, 0, 1);
            layout.Controls.Add(heading, 0, 0);

            var profile = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            profile.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 270));
            profile.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            profile.RowStyles.Add(new RowStyle(SizeType.Absolute, 31));
            profile.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _titleMode.DropDownStyle = ComboBoxStyle.DropDownList;
            _titleMode.AccessibleName = "南开草案排版模式";
            _titleMode.Dock = DockStyle.Fill;
            _titleMode.Items.AddRange(new object[] { "南开 2026 草案 · 模式一（数字标题）", "南开 2026 草案 · 模式二（中文序号）" });
            _titleMode.SelectedIndex = 0;
            _titleMode.SelectedIndexChanged += (sender, args) => InvalidateAcknowledgment();
            profile.Controls.Add(_titleMode, 0, 0);
            profile.Controls.Add(new Label
            {
                Text = "模式一：章/一级居左；模式二：章/一级居中、二/三级左缩进两字。仅排版，不自动改写标题编号。",
                Dock = DockStyle.Fill, Padding = new Padding(10, 5, 0, 0)
            }, 1, 0);
            var scope = new Label
            {
                Text = "草案范围：已覆盖角色的字体、字号、段落格式。未勾选项及未覆盖角色保留原格式；表格内部、图片内容、复杂公式、页眉页脚、目录域与参考文献内容不由此窗口改写。识别候选不等于校方认证。",
                Dock = DockStyle.Fill, Padding = new Padding(8, 5, 8, 4),
                BackColor = Color.FromArgb(250, 239, 216), ForeColor = Color.FromArgb(116, 74, 28)
            };
            profile.Controls.Add(scope, 0, 1);
            profile.SetColumnSpan(scope, 2);
            layout.Controls.Add(profile, 0, 1);

            var tools = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(0, 5, 0, 0)
            };
            _sectionFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            _sectionFilter.AccessibleName = "所属部分筛选";
            _sectionFilter.Width = 132;
            _sectionFilter.DisplayMember = nameof(Choice.Label);
            _sectionFilter.ValueMember = nameof(Choice.Value);
            _sectionFilter.Items.Add(new Choice("", "整篇 · 全部部分"));
            foreach (var item in SectionChoices()) _sectionFilter.Items.Add(item);
            _sectionFilter.SelectedIndex = 0;
            _sectionFilter.SelectedIndexChanged += (sender, args) => ApplyFilter();
            tools.Controls.Add(_sectionFilter);
            tools.Controls.Add(new Label { Text = "原文检索", AutoSize = true, Padding = new Padding(5, 5, 0, 0) });
            _search.Width = 158;
            _search.AccessibleName = "检索大纲原文";
            _search.TextChanged += (sender, args) => { _filterTimer.Stop(); _filterTimer.Start(); };
            _filterTimer.Tick += (sender, args) => { _filterTimer.Stop(); ApplyFilter(); };
            tools.Controls.Add(_search);
            var confirmRules = CreateButton("确认全篇高置信规则项", false);
            confirmRules.Click += (sender, args) => ConfirmRuleCandidates();
            tools.Controls.Add(confirmRules);
            _confirmVisible.Text = "确认此部分可见项";
            StyleButton(_confirmVisible, false);
            _confirmVisible.Click += (sender, args) => SetVisibleApprovals(true);
            tools.Controls.Add(_confirmVisible);
            var clear = CreateButton("取消可见确认", false);
            clear.Click += (sender, args) => SetVisibleApprovals(false);
            tools.Controls.Add(clear);
            layout.Controls.Add(tools, 0, 2);

            ConfigureGrid();
            layout.Controls.Add(_grid, 0, 3);
            _detail.Dock = DockStyle.Fill;
            _detail.AccessibleName = "完整原文与识别依据";
            _detail.Multiline = true;
            _detail.ReadOnly = true;
            _detail.ScrollBars = ScrollBars.Vertical;
            _detail.BackColor = Color.White;
            _detail.BorderStyle = BorderStyle.FixedSingle;
            _detail.Margin = new Padding(0, 8, 0, 0);
            _detail.Text = "选择一行查看完整原文与识别依据。";
            layout.Controls.Add(_detail, 0, 4);

            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3 };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 292));
            footer.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            footer.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            footer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _acknowledgment.Text = "我已核对勾选项，明确确认其结构角色，并选择使用上述 2026 草案范围。";
            _acknowledgment.Dock = DockStyle.Fill;
            _acknowledgment.CheckedChanged += (sender, args) => RefreshSummary();
            footer.Controls.Add(_acknowledgment, 0, 0);
            footer.SetColumnSpan(_acknowledgment, 2);
            _notice.Dock = DockStyle.Fill;
            _notice.ForeColor = Color.FromArgb(116, 74, 28);
            _notice.AutoEllipsis = true;
            _notice.Text = BuildWarning(review.Warnings);
            footer.Controls.Add(_notice, 0, 1);
            footer.SetColumnSpan(_notice, 2);
            _summary.Dock = DockStyle.Fill;
            _summary.Padding = new Padding(0, 6, 0, 0);
            _summary.ForeColor = Purple;
            footer.Controls.Add(_summary, 0, 2);
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false
            };
            var cancel = CreateButton("取消", false);
            cancel.DialogResult = DialogResult.Cancel;
            _confirm.Text = "确认大纲并继续";
            StyleButton(_confirm, true);
            _confirm.Click += (sender, args) => AcceptApprovals();
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_confirm);
            footer.Controls.Add(buttons, 1, 2);
            layout.Controls.Add(footer, 0, 5);
            Controls.Add(layout);
            AcceptButton = _confirm;
            CancelButton = cancel;

            LoadBlocks(review.Blocks);
            ApplyFilter();
        }

        private void ConfigureGrid()
        {
            _grid.Dock = DockStyle.Fill;
            _grid.Name = "ThesisOutlineGrid";
            _grid.AccessibleName = "整篇论文大纲";
            _grid.BackgroundColor = Color.White;
            _grid.BorderStyle = BorderStyle.FixedSingle;
            _grid.AllowUserToAddRows = _grid.AllowUserToDeleteRows = false;
            _grid.AllowUserToOrderColumns = false;
            _grid.AllowUserToResizeRows = false;
            _grid.AutoGenerateColumns = false;
            _grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            _grid.RowTemplate.Height = 43;
            _grid.RowHeadersVisible = false;
            _grid.MultiSelect = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.EditMode = DataGridViewEditMode.EditOnEnter;
            _grid.EnableHeadersVisualStyles = false;
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 231, 235);
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Purple;
            _grid.ColumnHeadersDefaultCellStyle.Font = OwnFont(9F, FontStyle.Bold);
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            _grid.ColumnHeadersHeight = 32;
            _grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(250, 239, 216);
            _grid.DefaultCellStyle.SelectionForeColor = ForeColor;
            _grid.DefaultCellStyle.Padding = new Padding(4, 2, 4, 2);
            _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(252, 250, 247);
            _grid.GridColor = Color.FromArgb(232, 224, 224);
            _grid.Columns.Add(new DataGridViewCheckBoxColumn
                { Name = ConfirmColumn, HeaderText = "确认", Width = 46, SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
                { Name = "paragraph", HeaderText = "段落", Width = 57, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(ChoiceColumn(SectionColumn, "所属部分", 100, SectionChoices()));
            _grid.Columns.Add(ChoiceColumn(KindColumn, "段落类型", 96, KindChoices()));
            _grid.Columns.Add(ChoiceColumn(LevelColumn, "标题层级", 100, LevelChoices()));
            _grid.Columns.Add(new DataGridViewTextBoxColumn
                { Name = "excerpt", HeaderText = "原文摘要", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                    MinimumWidth = 190, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
                { Name = StateColumn, HeaderText = "核对状态", Width = 150, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewButtonColumn
                { Name = "locate", HeaderText = "原文", Text = "定位", UseColumnTextForButtonValue = true,
                    Width = 58, SortMode = DataGridViewColumnSortMode.NotSortable, FlatStyle = FlatStyle.Flat });
            _grid.CurrentCellDirtyStateChanged += (sender, args) =>
            {
                if (_grid.IsCurrentCellDirty) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _grid.CellValueChanged += GridCellValueChanged;
            _grid.CellContentClick += (sender, args) =>
            {
                if (args.RowIndex < 0 || _grid.Columns[args.ColumnIndex].Name != "locate") return;
                var state = (RowState)_grid.Rows[args.RowIndex].Tag;
                LocateRequested?.Invoke(state.Source);
            };
            _grid.SelectionChanged += (sender, args) => ShowSelectedText();
            _grid.DataError += (sender, args) =>
            {
                args.ThrowException = false;
                _notice.Text = "此单元格的值无效，请从下拉列表中选择；无效项不会被批准。";
                if (args.RowIndex >= 0)
                {
                    var state = (RowState)_grid.Rows[args.RowIndex].Tag;
                    if (state != null) { state.Confirmed = false; RefreshRow(state); }
                }
                InvalidateAcknowledgment();
            };
        }

        private void LoadBlocks(IList<EngineStructureBlock> blocks)
        {
            _updating = true;
            _grid.SuspendLayout();
            try
            {
                foreach (var block in blocks)
                {
                    var state = new RowState(block);
                    int rowIndex = _grid.Rows.Add(false, (block.ParagraphIndex + 1).ToString(CultureInfo.InvariantCulture),
                        state.Section.ToString(), state.Kind.ToString(), LevelValue(state),
                        Excerpt(block.ExactText), string.Empty);
                    state.GridRow = _grid.Rows[rowIndex];
                    state.GridRow.Tag = state;
                    state.GridRow.Cells["excerpt"].ToolTipText = block.ExactText ?? string.Empty;
                    _rows.Add(state);
                    RefreshRow(state);
                }
            }
            finally { _grid.ResumeLayout(); _updating = false; }
        }

        private void GridCellValueChanged(object sender, DataGridViewCellEventArgs args)
        {
            if (_updating || args.RowIndex < 0 || args.ColumnIndex < 0) return;
            var state = (RowState)_grid.Rows[args.RowIndex].Tag;
            if (state == null) return;
            string column = _grid.Columns[args.ColumnIndex].Name;
            _updating = true;
            try
            {
                if (column == ConfirmColumn)
                    state.Confirmed = CanApprove(state) && Convert.ToBoolean(state.GridRow.Cells[ConfirmColumn].Value ?? false);
                else if (column == SectionColumn || column == KindColumn || column == LevelColumn)
                {
                    EngineThesisSection section;
                    EngineThesisBlockKind kind;
                    int level;
                    if (!Enum.TryParse(Convert.ToString(state.GridRow.Cells[SectionColumn].Value), out section) ||
                        !Enum.TryParse(Convert.ToString(state.GridRow.Cells[KindColumn].Value), out kind))
                        throw new InvalidOperationException("请选择有效的结构角色。");
                    state.Section = section;
                    state.Kind = kind;
                    int.TryParse(Convert.ToString(state.GridRow.Cells[LevelColumn].Value), out level);
                    state.HeadingLevel = kind == EngineThesisBlockKind.Heading ? (int?)(level > 0 ? level : 1) : null;
                    state.RoleEdited = true;
                    state.Confirmed = false;
                    state.GridRow.Cells[LevelColumn].Value = LevelValue(state);
                }
                RefreshRow(state);
            }
            catch (InvalidOperationException error)
            {
                state.Confirmed = false;
                _notice.Text = error.Message;
                RefreshRow(state);
            }
            finally { _updating = false; }
            InvalidateAcknowledgment();
            RefreshSummary();
        }

        private static bool CanApprove(RowState state)
        {
            return state.Source.CanConfirm && state.Source.HasCompleteTextMapping && !state.Source.IsTableCell &&
                state.Section != EngineThesisSection.Unknown && state.Kind != EngineThesisBlockKind.Unknown &&
                state.Kind != EngineThesisBlockKind.TableCell &&
                Enum.IsDefined(typeof(EngineThesisSection), state.Section) && Enum.IsDefined(typeof(EngineThesisBlockKind), state.Kind) &&
                (state.Kind != EngineThesisBlockKind.Heading || (state.HeadingLevel >= 1 && state.HeadingLevel <= 9));
        }

        private static bool HasFormatTarget(RowState state)
        {
            if (state.Kind == EngineThesisBlockKind.Heading)
                return state.Section == EngineThesisSection.Body
                    ? state.HeadingLevel >= 1 && state.HeadingLevel <= 4
                    : state.Section != EngineThesisSection.Unknown && state.Section != EngineThesisSection.FrontMatter && state.HeadingLevel == 1;
            if (state.Kind == EngineThesisBlockKind.Paragraph)
                return state.Section != EngineThesisSection.Unknown && state.Section != EngineThesisSection.FrontMatter && state.Section != EngineThesisSection.Contents;
            return (state.Section == EngineThesisSection.Body || state.Section == EngineThesisSection.Appendix) &&
                (state.Kind == EngineThesisBlockKind.FigureCaption || state.Kind == EngineThesisBlockKind.TableCaption);
        }

        private void RefreshRow(RowState state)
        {
            var row = state.GridRow;
            bool canApprove = CanApprove(state);
            if (!canApprove) state.Confirmed = false;
            row.Cells[ConfirmColumn].ReadOnly = !canApprove;
            row.Cells[ConfirmColumn].Value = state.Confirmed;
            row.Cells[LevelColumn].ReadOnly = state.Kind != EngineThesisBlockKind.Heading;
            string status;
            if (state.Source.IsTableCell || state.Kind == EngineThesisBlockKind.TableCell) status = "表格内部 · 保留";
            else if (!state.Source.CanConfirm || !state.Source.HasCompleteTextMapping) status = "锚点未支持 · 保留";
            else if (!canApprove) status = "角色未判定 · 请改正";
            else if (state.Confirmed) status = HasFormatTarget(state) ? "已明确确认" : "已确认 · 本版保留格式";
            else if (state.RoleEdited) status = "已修改角色 · 待确认";
            else if (state.Source.RequiresReview || !IsHighConfidence(state.Source)) status = "候选待核对 · 不自动批准";
            else status = "规则建议 · 待确认";
            row.Cells[StateColumn].Value = status;
            row.Cells[StateColumn].ToolTipText = (state.Source.Evidence ?? string.Empty) +
                (HasFormatTarget(state) ? string.Empty : "\r\n此角色没有本版格式目标，确认后仍保留原格式。");
            row.Cells[StateColumn].Style.ForeColor = state.Confirmed ? Purple : Color.FromArgb(116, 92, 68);
        }

        private static bool IsHighConfidence(EngineStructureBlock block) =>
            !double.IsNaN(block.Confidence) && !double.IsInfinity(block.Confidence) && block.Confidence >= 0.90 && block.Confidence <= 1;

        private void ConfirmRuleCandidates()
        {
            EndGridEdit();
            _updating = true;
            try
            {
                foreach (var state in _rows.Where(row => !row.RoleEdited && !row.Source.RequiresReview &&
                    IsHighConfidence(row.Source) && CanApprove(row) && HasFormatTarget(row)))
                { state.Confirmed = true; RefreshRow(state); }
            }
            finally { _updating = false; }
            _notice.Text = "已按你的操作确认全篇高置信规则项；需复核的候选没有被批量批准，请定位检查。";
            InvalidateAcknowledgment();
            RefreshSummary();
        }

        private void SetVisibleApprovals(bool confirm)
        {
            EndGridEdit();
            if (confirm && string.IsNullOrEmpty(SelectedSectionFilter())) return;
            _updating = true;
            try
            {
                foreach (var state in _rows.Where(row => row.GridRow.Visible && (!confirm || CanApprove(row))))
                { state.Confirmed = confirm; RefreshRow(state); }
            }
            finally { _updating = false; }
            _notice.Text = confirm
                ? "已明确确认当前部分中可见的已判定角色；未知角色、表格内部和不安全锚点仍不批准。"
                : "已取消可见项的确认。";
            InvalidateAcknowledgment();
            RefreshSummary();
        }

        private string SelectedSectionFilter() => (_sectionFilter.SelectedItem as Choice)?.Value ?? string.Empty;

        private void ApplyFilter()
        {
            if (_updating) return;
            EndGridEdit();
            string section = SelectedSectionFilter();
            string query = _search.Text.Trim();
            _grid.CurrentCell = null;
            _grid.SuspendLayout();
            try
            {
                foreach (var state in _rows)
                    state.GridRow.Visible = (section.Length == 0 || state.Section.ToString() == section) &&
                        (query.Length == 0 || (state.Source.ExactText ?? string.Empty).IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            finally { _grid.ResumeLayout(); }
            _confirmVisible.Enabled = section.Length > 0 && section != EngineThesisSection.Unknown.ToString();
            RefreshSummary();
        }

        private void ShowSelectedText()
        {
            if (_grid.CurrentRow == null) return;
            var state = _grid.CurrentRow.Tag as RowState;
            if (state == null) return;
            _detail.Text = "第 " + (state.Source.ParagraphIndex + 1) + " 段 · " + (state.Source.Location ?? string.Empty) +
                "\r\n" + (state.Source.ExactText ?? string.Empty) + "\r\n\r\n识别依据：" + (state.Source.Evidence ?? string.Empty);
        }

        private void RefreshSummary()
        {
            int count = _rows.Count(row => row.Confirmed && CanApprove(row));
            int visible = _rows.Count(row => row.GridRow.Visible);
            _summary.Text = "整篇 " + _rows.Count + " 段 · 可见 " + visible + " 段 · 已明确确认 " + count + " 段";
            _confirm.Enabled = count > 0 && _acknowledgment.Checked;
        }

        private void InvalidateAcknowledgment()
        {
            _acknowledgment.Checked = false;
            _confirm.Enabled = false;
        }

        private void EndGridEdit()
        {
            if (_grid.IsCurrentCellDirty) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            _grid.EndEdit();
        }

        private void AcceptApprovals()
        {
            EndGridEdit();
            if (!_acknowledgment.Checked) return;
            var approved = _rows.Where(row => row.Confirmed && CanApprove(row)).Select(row => new EngineApprovedStructureBlock
            {
                BlockId = row.Source.BlockId, Section = row.Section, Kind = row.Kind,
                HeadingLevel = row.Kind == EngineThesisBlockKind.Heading ? row.HeadingLevel : null
            }).ToList();
            if (approved.Count == 0) return;
            ApprovedBlocks = approved;
            DialogResult = DialogResult.OK;
            Close();
        }

        internal DialogResult ShowForOfficeWindow(int windowHandle)
        {
            var handle = new IntPtr(windowHandle);
            if (!_scaledForOffice)
            {
                uint dpi = 96;
                try { dpi = GetDpiForWindow(handle); } catch (EntryPointNotFoundException) { }
                float scale = Math.Max(1F, Math.Min(4F, dpi / 96F));
                var controls = new List<Control>();
                Collect(this, controls);
                Font[] fonts = controls.Select(control => control.Font).ToArray();
                Font headerFont = _grid.ColumnHeadersDefaultCellStyle.Font;
                Scale(new SizeF(scale, scale));
                for (int index = 0; index < controls.Count; index++)
                    controls[index].Font = OwnScaledFont(fonts[index], scale);
                _grid.ColumnHeadersDefaultCellStyle.Font = OwnScaledFont(headerFont, scale);
                _grid.RowTemplate.Height = (int)Math.Round(43 * scale);
                _grid.ColumnHeadersHeight = (int)Math.Round(32 * scale);
                foreach (DataGridViewRow row in _grid.Rows) row.Height = (int)Math.Round(43 * scale);
                foreach (DataGridViewColumn column in _grid.Columns)
                {
                    column.MinimumWidth = Math.Max(5, (int)Math.Round(column.MinimumWidth * scale));
                    if (column.AutoSizeMode != DataGridViewAutoSizeColumnMode.Fill)
                        column.Width = (int)Math.Round(column.Width * scale);
                }
                Rectangle work = Screen.FromHandle(handle).WorkingArea;
                int maxWidth = Math.Max(1, work.Width - 36);
                int maxHeight = Math.Max(1, work.Height - 36);
                MinimumSize = new Size(Math.Min((int)Math.Round(900 * scale), maxWidth),
                    Math.Min((int)Math.Round(620 * scale), maxHeight));
                Size = new Size(Math.Min(Width, maxWidth), Math.Min(Height, maxHeight));
                StartPosition = FormStartPosition.Manual;
                Location = new Point(work.Left + (work.Width - Width) / 2, work.Top + (work.Height - Height) / 2);
                _scaledForOffice = true;
            }
            return ShowDialog(new OfficeOwner(windowHandle));
        }

        private Font OwnFont(float points, FontStyle style = FontStyle.Regular)
        {
            var font = new Font("Microsoft YaHei UI", points, style);
            _ownedFonts.Add(font);
            return font;
        }

        private Font OwnScaledFont(Font original, float scale)
        {
            var font = new Font(original.FontFamily, original.SizeInPoints * 96F / 72F * scale, original.Style, GraphicsUnit.Pixel);
            _ownedFonts.Add(font);
            return font;
        }

        private static DataGridViewComboBoxColumn ChoiceColumn(string name, string label, int width, Choice[] values) =>
            new DataGridViewComboBoxColumn
            {
                Name = name, HeaderText = label, Width = width, DataSource = values,
                DisplayMember = nameof(Choice.Label), ValueMember = nameof(Choice.Value),
                FlatStyle = FlatStyle.Flat, DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };

        private Button CreateButton(string text, bool primary)
        {
            var button = new Button { Text = text };
            StyleButton(button, primary);
            return button;
        }

        private static void StyleButton(Button button, bool primary)
        {
            button.AutoSize = true;
            button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            button.Padding = new Padding(8, 2, 8, 2);
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = primary ? Purple : Color.White;
            button.ForeColor = primary ? Color.White : Purple;
            button.FlatAppearance.BorderColor = Color.FromArgb(208, 187, 196);
            button.FlatAppearance.BorderSize = primary ? 0 : 1;
            button.UseVisualStyleBackColor = false;
            button.Cursor = Cursors.Hand;
        }

        private static Choice[] SectionChoices() => new[]
        {
            new Choice("Unknown", "未判定"), new Choice("FrontMatter", "封面与前置部分"),
            new Choice("ChineseAbstract", "中文摘要"), new Choice("EnglishAbstract", "英文摘要"),
            new Choice("Contents", "目录"), new Choice("Body", "正文"), new Choice("References", "参考文献"),
            new Choice("Appendix", "附录"), new Choice("Acknowledgments", "致谢")
        };

        private static Choice[] KindChoices() => new[]
        {
            new Choice("Unknown", "未判定"), new Choice("Paragraph", "段落文字"), new Choice("Heading", "标题"),
            new Choice("Keywords", "关键词"), new Choice("ContentsEntry", "目录条目"),
            new Choice("FigureCaption", "图题"), new Choice("TableCaption", "表题"), new Choice("TableCell", "表格单元格")
        };

        private static Choice[] LevelChoices() => new[]
        {
            new Choice("0", "不适用"), new Choice("1", "章 / 部分标题"), new Choice("2", "一级标题"),
            new Choice("3", "二级标题"), new Choice("4", "三级标题"), new Choice("5", "第 5 层（保留）"),
            new Choice("6", "第 6 层（保留）"), new Choice("7", "第 7 层（保留）"),
            new Choice("8", "第 8 层（保留）"), new Choice("9", "第 9 层（保留）")
        };

        private static string LevelValue(RowState state) => state.Kind == EngineThesisBlockKind.Heading
            ? (state.HeadingLevel ?? 1).ToString(CultureInfo.InvariantCulture) : "0";
        private static string Excerpt(string text)
        {
            string line = (text ?? string.Empty).Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
            return line.Length > 110 ? line.Substring(0, 110) + "…" : line;
        }
        private static string BuildWarning(IList<string> warnings) => warnings != null && warnings.Count > 0
            ? "识别提示：" + string.Join("；", warnings.Take(2))
            : "尚未确认任何候选。可按所属部分筛选后批量确认已核对项；未知角色和表格内部不批准。";
        private static void Collect(Control control, List<Control> result)
        { result.Add(control); foreach (Control child in control.Controls) Collect(child, result); }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _filterTimer.Dispose();
            base.Dispose(disposing);
            if (disposing) foreach (Font font in _ownedFonts) font.Dispose();
        }

        private sealed class RowState
        {
            internal readonly EngineStructureBlock Source;
            internal EngineThesisSection Section;
            internal EngineThesisBlockKind Kind;
            internal int? HeadingLevel;
            internal bool Confirmed;
            internal bool RoleEdited;
            internal DataGridViewRow GridRow;
            internal RowState(EngineStructureBlock source)
            {
                Source = source; Section = source.Section; Kind = source.Kind;
                HeadingLevel = source.Kind == EngineThesisBlockKind.Heading ? source.HeadingLevel ?? 1 : null;
            }
        }

        private sealed class Choice
        {
            public string Value { get; }
            public string Label { get; }
            internal Choice(string value, string label) { Value = value; Label = label; }
            public override string ToString() => Label;
        }

        private sealed class OfficeOwner : IWin32Window
        { public IntPtr Handle { get; } internal OfficeOwner(int handle) { Handle = new IntPtr(handle); } }
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
    }
}
