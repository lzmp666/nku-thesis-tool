using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>
    /// 修复预览：在更新当前文档之前，逐条展示会改动的位置、改前值、目标值和依据，
    /// 并单独列出保留给用户的诊断。用户确认后才执行。
    /// </summary>
    public sealed class FixPreviewDialog : Form
    {
        public FixPreviewDialog(EnginePatchPlan plan, string scopeLabel)
        {
            Text = "修复预览（" + scopeLabel + "）";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(760, 500);
            MinimizeBox = false;
            MaximizeBox = false;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(246, 248, 251);

            var summary = new Label
            {
                Dock = DockStyle.Top,
                Height = 70,
                Padding = new Padding(12, 8, 10, 8),
                BackColor = Color.FromArgb(255, 247, 220),
                ForeColor = Color.FromArgb(91, 63, 12),
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                Text = "演示规则预览 · 将修改 " + plan.Patches.Count + " 处，"
                    + "保留 " + plan.Deferred.Count + " 项给用户处理。"
                    + Environment.NewLine
                    + "确认后直接更新当前打开的文档；不自动保存，是否保存由你决定。"
            };

            var list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                GridLines = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9F),
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            list.Columns.Add("格式项目", 100);
            list.Columns.Add("位置", 220);
            list.Columns.Add("当前值", 110);
            list.Columns.Add("修改为", 110);
            list.Columns.Add("修改说明", 180);
            foreach (var group in plan.Patches.GroupBy(patch => new
                     {
                patch.RuleId,
                         patch.DesiredAfter,
                         patch.Reason
                     }))
            {
                EnginePatch patch = group.First();
                var item = new ListViewItem(DiagnosticsPane.DescribeRule(patch.RuleId));
                item.SubItems.Add(group.Count() == 1
                    ? patch.Location
                    : group.Count() + " 处（" + patch.Location + " 等）");
                item.SubItems.Add(string.Join("、", group.Select(value => value.ExpectedBefore)
                    .Distinct(StringComparer.Ordinal)));
                item.SubItems.Add(patch.DesiredAfter);
                item.SubItems.Add(patch.Reason);
                list.Items.Add(item);
            }

            var deferred = new TextBox
            {
                Dock = DockStyle.Bottom,
                Height = 100,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 8.5F),
                Text = BuildDeferredText(plan)
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 54,
                Padding = new Padding(8),
                FlowDirection = FlowDirection.RightToLeft,
                BackColor = BackColor
            };
            var cancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                Padding = new Padding(10, 3, 10, 3),
                BackColor = Color.White
            };
            cancel.FlatAppearance.BorderColor = Color.FromArgb(197, 207, 221);
            var confirm = new Button
            {
                Text = "确认并修复当前文档",
                DialogResult = DialogResult.OK,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                Padding = new Padding(10, 3, 10, 3),
                BackColor = Color.FromArgb(37, 99, 190),
                ForeColor = Color.White
            };
            confirm.FlatAppearance.BorderSize = 0;
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(confirm);

            Controls.Add(list);
            Controls.Add(deferred);
            Controls.Add(buttons);
            Controls.Add(summary);
            AcceptButton = confirm;
            CancelButton = cancel;
        }

        private static string BuildDeferredText(EnginePatchPlan plan)
        {
            if (plan.Deferred.Count == 0)
            {
                return "保留项：无。";
            }

            var lines = new System.Collections.Generic.List<string> { "以下项目不会自动修改，需要你确认或手动处理：" };
            foreach (EngineDiagnostic diagnostic in plan.Deferred)
            {
                int occurrenceCount = diagnostic.Occurrences != null && diagnostic.Occurrences.Count > 0
                    ? diagnostic.Occurrences.Count : 1;
                lines.Add("· " + DiagnosticsPane.DescribeRule(diagnostic.RuleId) + " @ "
                    + (occurrenceCount > 1 ? occurrenceCount + " 处（" + diagnostic.Location + " 等）" : diagnostic.Location)
                    + "（处理方式：" + DiagnosticsPane.DescribeFixLevel(diagnostic.FixLevel)
                    + "；期望 " + diagnostic.Expected + "，实际 " + diagnostic.Actual + "）");
            }

            return string.Join(Environment.NewLine, lines.ToArray());
        }
    }
}
