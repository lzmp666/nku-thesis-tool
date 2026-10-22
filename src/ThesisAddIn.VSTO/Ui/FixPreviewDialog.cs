using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>
    /// 修复预览：在更新当前工作副本之前，逐条展示会改动的位置、改前值、目标值和依据，
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
                    + "保留 " + plan.Deferred.Count + " 项给用户处理。"
                    + Environment.NewLine
                    + "修复直接更新当前标注工作副本；原始提交文件不变，完成后由你决定是否保存。"
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
            foreach (var group in plan.Patches.GroupBy(patch => new
                     {
                         patch.RuleId,
                         patch.DesiredAfter,
                         patch.Reason
                     }))
            {
                EnginePatch patch = group.First();
                var item = new ListViewItem(patch.RuleId);
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
            var confirm = new Button { Text = "确认并修复当前工作副本", DialogResult = DialogResult.OK, AutoSize = true };
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
                lines.Add("· " + diagnostic.RuleId + " @ "
                    + (occurrenceCount > 1 ? occurrenceCount + " 处（" + diagnostic.Location + " 等）" : diagnostic.Location)
                    + "（处理方式：" + DiagnosticsPane.DescribeFixLevel(diagnostic.FixLevel)
                    + "；期望 " + diagnostic.Expected + "，实际 " + diagnostic.Actual + "）");
            }

            return string.Join(Environment.NewLine, lines.ToArray());
        }
    }
}
