using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ThesisAddIn.VSTO.Engine;

namespace ThesisAddIn.VSTO.Ui
{
    internal sealed class ObjectFormattingDialog : Form
    {
        internal readonly ObjectFormattingSettings Settings = new ObjectFormattingSettings();
        private readonly List<Font> _displayFonts = new List<Font>();
        internal ObjectFormattingDialog()
        {
            Text = "图表排版 · 当前文档";
            AutoScaleMode = AutoScaleMode.None;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(250, 247, 241);
            ForeColor = Color.FromArgb(44, 39, 42);
            ClientSize = new Size(570, 500);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            var layout = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
                WrapContents = false, Padding = new Padding(22), AutoScroll = true };
            layout.Controls.Add(new Label { Text = "南开 · 图表排版", Font = new Font(Font.FontFamily, 15, FontStyle.Bold),
                ForeColor = Color.FromArgb(104,31,61), AutoSize = true, Margin = new Padding(0,0,0,12) });
            layout.Controls.Add(new Label { Text = "适用于与题注配对的行内图片及简单表格。只修改当前文档，完成后自行保存。\r\n表内字号与线宽为作者排版选项，不是校级强制数值；复杂对象会列明跳过原因。",
                AutoSize = true, MaximumSize = new Size(520, 0), Margin = new Padding(0,0,0,12) });
            Add(layout, "图片居中（保留图片内容）", value => Settings.CenterPictures = value);
            Add(layout, "超宽图片等比缩小到当前版心", value => Settings.FitPictures = value);
            Add(layout, "表格整体居中；超宽固定表格按原列宽比例缩小", value => Settings.CenterTables = Settings.FitTables = value);
            Add(layout, "表内中文宋体，英文 Times New Roman", value => Settings.TableFont = value);
            var sizeRow = new FlowLayoutPanel { Width = 510, Height = 34, Margin = new Padding(0) };
            sizeRow.Controls.Add(new Label { Text = "表内字号（磅）：", AutoSize = true, Padding = new Padding(0,5,0,0) });
            var size = new NumericUpDown { Minimum = 8, Maximum = 24, DecimalPlaces = 1, Increment = 0.5M,
                Value = 10.5M, Width = 75 };
            size.ValueChanged += (s,e) => Settings.TableSize = (float)size.Value;
            sizeRow.Controls.Add(size); layout.Controls.Add(sizeRow);
            Add(layout, "表内文字水平、垂直居中", value => Settings.TableAlignment = value);
            Add(layout, "三线表：顶底线 1.5 磅、表头下线 0.75 磅，去除竖线", value => Settings.ThreeLine = value);
            Add(layout, "首行为表头，跨页重复首行", value => Settings.RepeatHeader = value);
            Add(layout, "题注按已核对的规范设置字体、字号、对齐与段落间距", value => Settings.Captions = value);
            var buttons = new FlowLayoutPanel { Width = 510, Height = 44, FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0,12,0,0) };
            var apply = new Button { Text = "应用到配对图表", DialogResult = DialogResult.OK,
                Width = 165, Height = 34, BackColor = Color.FromArgb(104,31,61), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 90, Height = 34 };
            buttons.Controls.Add(apply); buttons.Controls.Add(cancel); layout.Controls.Add(buttons);
            Controls.Add(layout); AcceptButton = apply; CancelButton = cancel;
        }
        private static void Add(Control parent, string text, Action<bool> set)
        {
            var check = new CheckBox { Text = text, Checked = true, AutoSize = true,
                Margin = new Padding(0,3,0,3) };
            check.CheckedChanged += (s,e) => set(check.Checked); parent.Controls.Add(check);
        }
        internal DialogResult ShowForOfficeWindow(int handle)
        {
            uint dpi = 96;
            try { dpi = GetDpiForWindow(new IntPtr(handle)); } catch (EntryPointNotFoundException) { }
            float scale = Math.Max(1F, Math.Min(4F, dpi / 96F));
            var controls = new List<Control>(); Collect(this, controls);
            Font[] fonts = controls.Select(control => control.Font).ToArray();
            Scale(new SizeF(scale, scale));
            for (int index = 0; index < controls.Count; index++) {
                var font = new Font(fonts[index].FontFamily, fonts[index].SizeInPoints * 96F / 72F * scale,
                    fonts[index].Style, GraphicsUnit.Pixel);
                _displayFonts.Add(font); controls[index].Font = font;
            }
            return ShowDialog(new OfficeOwner(handle));
        }
        private static void Collect(Control control, List<Control> result)
        { result.Add(control); foreach (Control child in control.Controls) Collect(child, result); }
        private sealed class OfficeOwner : IWin32Window
        { public IntPtr Handle { get; } internal OfficeOwner(int handle) { Handle = new IntPtr(handle); } }
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
        protected override void Dispose(bool disposing)
        { base.Dispose(disposing); if (disposing) foreach (Font font in _displayFonts) font.Dispose(); }
    }
}
