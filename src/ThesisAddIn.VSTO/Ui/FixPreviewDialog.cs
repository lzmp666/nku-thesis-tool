using System;
using System.Drawing;
using System.Windows.Forms;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>
    /// 修复预览：在真正写入新文件之前，逐条展示会改动的位置、改前值、目标值和依据，
    /// 并单独列出保留给用户的诊断。用户确认后才执行。
    /// </summary>
    public sealed class FixPreviewDialog : Form
    {
        public FixPreviewDialog(EnginePatchPlan plan, string scopeLabel)
        {
            Text = "修复预览（" + scopeLabel + "）";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(720, 460);
            MinimizeBox = false;
            MaximizeBox = false;
            Font = new Font("Microsoft YaHei UI", 9F);

            var summary = new Label
            {
                Dock = DockStyle.Top,
                Height = 58,
                Padding = new Padding(8),
                Text = "将修改 " + plan.Patches.Count + " 处，"
                    + "保留 " + plan.Deferred.Count + " 处给用户处理。"
                    + Environment.NewLine
                    + "原始文件不会被修改；修复结果会另存为你指定的新文件。"
            };

            var list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false
            };
            list.Columns.Add("规则", 150);
            list.Columns.Add("位置", 170);
            list.Columns.Add("改前", 110);
            list.Columns.Add("改为", 110);
            list.Columns.Add("依据", 200);
            foreach (EnginePatch patch in plan.Patches)
            {
                var item = new ListViewItem(patch.RuleId);
                item.SubItems.Add(patch.Location);
                item.SubItems.Add(patch.ExpectedBefore);
                item.SubItems.Add(patch.DesiredAfter);
                item.SubItems.Add(patch.Reason);
                list.Items.Add(item);
            }

            var deferred = new TextBox
            {
                Dock = DockStyle.Bottom,
                Height = 84,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Text = BuildDeferredText(plan)
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                FlowDirection = FlowDirection.RightToLeft
            };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
            var confirm = new Button { Text = "确认修复并另存", DialogResult = DialogResult.OK, AutoSize = true };
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
                lines.Add("· " + diagnostic.RuleId + " @ " + diagnostic.Location
                    + "（处理方式：" + DiagnosticsPane.DescribeFixLevel(diagnostic.FixLevel)
                    + "；期望 " + diagnostic.Expected + "，实际 " + diagnostic.Actual + "）");
            }

            return string.Join(Environment.NewLine, lines.ToArray());
        }
    }
}
