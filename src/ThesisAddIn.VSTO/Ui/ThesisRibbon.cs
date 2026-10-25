using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>把分析、定位和修复操作放在 Word 功能区；侧边栏仅用于检索问题。</summary>
    [ComVisible(true)]
    public sealed class ThesisRibbon : Office.IRibbonExtensibility
    {
        private readonly Action _analyze;
        private readonly Action _fixCurrent;
        private readonly Action _fixAll;
        private readonly Action _verifyResolved;
        private readonly Action<int> _navigateIssue;
        private readonly Action<int> _navigateOccurrence;
        private readonly Action<bool> _setIssuesVisible;
        private readonly Func<bool> _isIssuesVisible;
        private readonly Func<bool> _canFixCurrent;
        private readonly Func<bool> _canFixAll;
        private readonly Func<bool> _canVerifyResolved;
        private readonly Func<bool> _canNavigateIssue;
        private readonly Func<bool> _canNavigateOccurrence;
        private Office.IRibbonUI _ribbonUi;
        private int _stateMask = -1;
        internal Action FormatObjects { get; set; }
        internal Action UndoBatch { get; set; }
        internal Func<bool> CanUndoBatch { get; set; }
        private readonly Dictionary<string, stdole.IPictureDisp> _images =
            new Dictionary<string, stdole.IPictureDisp>(StringComparer.Ordinal);

        public ThesisRibbon(Action analyze, Action fixCurrent, Action fixAll, Action verifyResolved,
            Action<int> navigateIssue, Action<int> navigateOccurrence, Action<bool> setIssuesVisible,
            Func<bool> isIssuesVisible,
            Func<bool> canFixCurrent, Func<bool> canFixAll, Func<bool> canVerifyResolved,
            Func<bool> canNavigateIssue, Func<bool> canNavigateOccurrence)
        {
            _analyze = analyze ?? throw new ArgumentNullException(nameof(analyze));
            _fixCurrent = fixCurrent ?? throw new ArgumentNullException(nameof(fixCurrent));
            _fixAll = fixAll ?? throw new ArgumentNullException(nameof(fixAll));
            _verifyResolved = verifyResolved ?? throw new ArgumentNullException(nameof(verifyResolved));
            _navigateIssue = navigateIssue ?? throw new ArgumentNullException(nameof(navigateIssue));
            _navigateOccurrence = navigateOccurrence ?? throw new ArgumentNullException(nameof(navigateOccurrence));
            _setIssuesVisible = setIssuesVisible ?? throw new ArgumentNullException(nameof(setIssuesVisible));
            _isIssuesVisible = isIssuesVisible ?? throw new ArgumentNullException(nameof(isIssuesVisible));
            _canFixCurrent = canFixCurrent ?? throw new ArgumentNullException(nameof(canFixCurrent));
            _canFixAll = canFixAll ?? throw new ArgumentNullException(nameof(canFixAll));
            _canVerifyResolved = canVerifyResolved ?? throw new ArgumentNullException(nameof(canVerifyResolved));
            _canNavigateIssue = canNavigateIssue ?? throw new ArgumentNullException(nameof(canNavigateIssue));
            _canNavigateOccurrence = canNavigateOccurrence ?? throw new ArgumentNullException(nameof(canNavigateOccurrence));
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""NkuThesisTab"" label=""论文格式检查"" insertAfterMso=""TabHome"">
        <group id=""NkuThesisIdentity"" label=""南开大学"">
          <button id=""NkuEmblem"" label=""南开大学"" size=""large""
                  getImage=""GetBrandImage"" onAction=""OnBrandInfo""
                  screentip=""NKU Thesis Tool"" supertip=""查看插件信息与南开校园主题。"" />
        </group>
        <group id=""NkuThesisCheck"" label=""检查与检索"">
          <button id=""NkuAnalyzeCurrent"" label=""分析当前文档"" size=""large"" getImage=""GetButtonImage""
                  screentip=""分析当前文档""
                  supertip=""分析当前 Word 文档，并在本页更新格式问题标注。""
                  onAction=""OnAnalyzeCurrent"" />
          <toggleButton id=""NkuShowIssues"" getLabel=""GetIssuesLabel"" size=""large""
                        getImage=""GetButtonImage"" getPressed=""GetIssuesPressed""
                        screentip=""打开或关闭问题检索列表""
                        supertip=""在当前 Word 窗口显示或隐藏左侧问题检索列表。""
                        onAction=""OnToggleIssues"" />
        </group>
        <group id=""NkuThesisRepair"" label=""修复"">
          <button id=""NkuFixCurrent"" label=""修复当前问题"" size=""large"" getImage=""GetButtonImage"" getEnabled=""GetEnabled""
                  screentip=""直接解决当前格式问题""
                  supertip=""按当前规则直接修改所选项，可在问题检索中撤销；不会自动保存。""
                  onAction=""OnFixCurrent"" />
          <button id=""NkuFixAll"" label=""批量修复"" size=""large"" getImage=""GetButtonImage"" getEnabled=""GetEnabled""
                  screentip=""选择需要修复的问题""
                  supertip=""进入勾选模式，在检索列表中选择后修复所选项。""
                  onAction=""OnFixAll"" />
          <button id=""NkuVerifyResolved"" label=""撤销当前修复"" getImage=""GetButtonImage"" getEnabled=""GetEnabled""
                  screentip=""撤销这一项的格式修复""
                  supertip=""仅恢复该项修复前的格式，保留其他修复和文字编辑；格式已被手动更改时停止。""
                  onAction=""OnVerifyResolved"" />
          <button id=""NkuFormatObjects"" label=""图表排版"" getImage=""GetButtonImage""
                  screentip=""图片与表格排版选项""
                  supertip=""为配对图表设置居中、限宽、表内中英文字体字号、单元格对齐、三线表及重复表头；作者选项不作为校级合规结论。""
                  onAction=""OnFormatObjects"" />
          <button id=""NkuUndoBatch"" label=""批量撤销"" getImage=""GetButtonImage"" getEnabled=""GetEnabled""
                  screentip=""勾选并撤销已修复项"" supertip=""在左侧勾选已修复项，再点撤销所选；保留其他修复和文字编辑。""
                  onAction=""OnUndoBatch"" />
        </group>
        <group id=""NkuThesisNavigation"" label=""定位"">
          <box id=""NkuIssueNavigation"" boxStyle=""vertical"">
            <button id=""NkuPreviousIssue"" label=""上一项"" getImage=""GetButtonImage"" getEnabled=""GetEnabled"" onAction=""OnPreviousIssue"" />
            <button id=""NkuNextIssue"" label=""下一项"" getImage=""GetButtonImage"" getEnabled=""GetEnabled"" onAction=""OnNextIssue"" />
          </box>
          <box id=""NkuOccurrenceNavigation"" boxStyle=""vertical"">
            <button id=""NkuPreviousOccurrence"" label=""上一处"" getImage=""GetButtonImage"" getEnabled=""GetEnabled"" onAction=""OnPreviousOccurrence"" />
            <button id=""NkuNextOccurrence"" label=""下一处"" getImage=""GetButtonImage"" getEnabled=""GetEnabled"" onAction=""OnNextOccurrence"" />
          </box>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUi)
        {
            _ribbonUi = ribbonUi;
            Invalidate();
        }

        public void Invalidate()
        {
            int state = (SafeGet(_canFixCurrent) ? 1 : 0) | (SafeGet(_canFixAll) ? 2 : 0)
                | (SafeGet(_canVerifyResolved) ? 4 : 0) | (SafeGet(_canNavigateIssue) ? 8 : 0)
                | (SafeGet(_canNavigateOccurrence) ? 16 : 0) | (SafeGet(_isIssuesVisible) ? 32 : 0)
                | (SafeGet(CanUndoBatch) ? 64 : 0);
            int changed = _stateMask < 0 ? 127 : _stateMask ^ state;
            _stateMask = state;
            if (_ribbonUi == null || changed == 0) return;
            try
            {
                // Selection changes must not rebuild the whole native Ribbon, which
                // also emits accessibility tab events and can relayout the task pane.
                if ((changed & 1) != 0) _ribbonUi.InvalidateControl("NkuFixCurrent");
                if ((changed & 2) != 0) _ribbonUi.InvalidateControl("NkuFixAll");
                if ((changed & 4) != 0) _ribbonUi.InvalidateControl("NkuVerifyResolved");
                if ((changed & 8) != 0)
                {
                    _ribbonUi.InvalidateControl("NkuPreviousIssue");
                    _ribbonUi.InvalidateControl("NkuNextIssue");
                }
                if ((changed & 16) != 0)
                {
                    _ribbonUi.InvalidateControl("NkuPreviousOccurrence");
                    _ribbonUi.InvalidateControl("NkuNextOccurrence");
                }
                if ((changed & 32) != 0) _ribbonUi.InvalidateControl("NkuShowIssues");
                if ((changed & 64) != 0) _ribbonUi.InvalidateControl("NkuUndoBatch");
            }
            catch (COMException) { }
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "NkuFixCurrent": return (_stateMask & 1) != 0;
                case "NkuFixAll": return (_stateMask & 2) != 0;
                case "NkuVerifyResolved": return (_stateMask & 4) != 0;
                case "NkuUndoBatch": return (_stateMask & 64) != 0;
                case "NkuPreviousIssue":
                case "NkuNextIssue": return (_stateMask & 8) != 0;
                case "NkuPreviousOccurrence":
                case "NkuNextOccurrence": return (_stateMask & 16) != 0;
                default: return true;
            }
        }

        public bool GetIssuesPressed(Office.IRibbonControl control) => (_stateMask & 32) != 0;
        public string GetIssuesLabel(Office.IRibbonControl control) =>
            GetIssuesPressed(control) ? "关闭检索" : "打开检索";

        public stdole.IPictureDisp GetButtonImage(Office.IRibbonControl control)
        {
            if (control == null) return null;
            if (_images.TryGetValue(control.Id, out stdole.IPictureDisp cached)) return cached;
            int index;
            switch (control.Id)
            {
                case "NkuAnalyzeCurrent": index = 0; break;
                case "NkuShowIssues": index = 1; break;
                case "NkuFixCurrent": index = 2; break;
                case "NkuFixAll": index = 3; break;
                case "NkuFormatObjects": index = 3; break;
                case "NkuVerifyResolved": index = 4; break;
                case "NkuUndoBatch": index = 4; break;
                case "NkuPreviousIssue": index = 5; break;
                case "NkuNextIssue": index = 6; break;
                case "NkuPreviousOccurrence": index = 7; break;
                case "NkuNextOccurrence": index = 8; break;
                default: return null;
            }

            using (Stream stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("ThesisAddIn.VSTO.Assets.nku-ribbon-icons.png"))
            {
                if (stream == null) return null;
                using (Image sheet = Image.FromStream(stream))
                using (var icon = new Bitmap(32, 32))
                using (Graphics graphics = Graphics.FromImage(icon))
                {
                    int column = index % 3;
                    int row = index / 3;
                    var source = Rectangle.FromLTRB(column * sheet.Width / 3, row * sheet.Height / 3,
                        (column + 1) * sheet.Width / 3, (row + 1) * sheet.Height / 3);
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.DrawImage(sheet, new Rectangle(0, 0, 32, 32), source, GraphicsUnit.Pixel);
                    stdole.IPictureDisp picture = RibbonPictureConverter.Convert(icon);
                    _images[control.Id] = picture;
                    return picture;
                }
            }
        }

        public stdole.IPictureDisp GetBrandImage(Office.IRibbonControl control)
        {
            if (control == null) return null;
            if (_images.TryGetValue(control.Id, out stdole.IPictureDisp cached)) return cached;
            bool emblem = control.Id == "NkuEmblem";
            string resource = emblem ? "nku-emblem-official.jpg" : "nku-campus-panorama-v2.png";
            using (Stream stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("ThesisAddIn.VSTO.Assets." + resource))
            {
                if (stream == null) return null;
                using (Image source = Image.FromStream(stream))
                using (var bitmap = new Bitmap(emblem ? 32 : 96, 32))
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.DrawImage(source, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                    cached = RibbonPictureConverter.Convert(bitmap);
                    _images[control.Id] = cached;
                    return cached;
                }
            }
        }

        public void OnBrandInfo(Office.IRibbonControl control)
        {
            WordWindowContext context = WordWindowContext.FromControl(control);
            uint dpi = 96;
            try { if (context != null) dpi = GetDpiForWindow(context.Handle); }
            catch (EntryPointNotFoundException) { }
            float scale = Math.Max(1F, Math.Min(4F, dpi / 96F));
            using (var dialog = new Form
            {
                Text = "NKU Thesis Tool · 南开校园主题", Width = (int)(760 * scale), Height = (int)(380 * scale),
                AutoScaleMode = AutoScaleMode.None, StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false,
                BackColor = Color.FromArgb(250, 247, 241)
            })
            using (Stream stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("ThesisAddIn.VSTO.Assets.nku-campus-panorama-v2.png"))
            using (Image illustration = stream == null ? null : Image.FromStream(stream))
            {
                var art = new PictureBox { Dock = DockStyle.Fill, Image = illustration, SizeMode = PictureBoxSizeMode.Zoom };
                var caption = new Label { Dock = DockStyle.Bottom, Height = (int)(76 * scale), TextAlign = ContentAlignment.MiddleCenter,
                    ForeColor = Color.FromArgb(118, 34, 85), Font = new Font("Microsoft YaHei UI", 11F * 96F / 72F * scale, GraphicsUnit.Pixel),
                    Text = GetBuildLabel(control) + "\r\n允公允能 · 日新月异\r\n校徽取自南开大学官网；校园主题插画由 GPT-image 生成。" };
                dialog.Controls.Add(art);
                dialog.Controls.Add(caption);
                if (context == null) dialog.ShowDialog(); else dialog.ShowDialog(context);
            }
        }

        public string GetBuildLabel(Office.IRibbonControl control)
        {
            string version = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "版本未知";
            return "预览版 " + version.Replace("-preview.", " · ");
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr window);

        private sealed class WordWindowContext : IWin32Window
        {
            private WordWindowContext(int handle) { Handle = new IntPtr(handle); }
            public IntPtr Handle { get; }
            internal static WordWindowContext FromControl(Office.IRibbonControl control)
            {
                try
                {
                    var window = control?.Context as Microsoft.Office.Interop.Word.Window;
                    return window == null ? null : new WordWindowContext(window.Hwnd);
                }
                catch (COMException) { return null; }
            }
        }

        private sealed class RibbonPictureConverter : AxHost
        {
            private RibbonPictureConverter() : base(string.Empty) { }
            internal static stdole.IPictureDisp Convert(Image image) =>
                (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }

        public void OnAnalyzeCurrent(Office.IRibbonControl control) => _analyze();
        public void OnFixCurrent(Office.IRibbonControl control) => _fixCurrent();
        public void OnFixAll(Office.IRibbonControl control) => _fixAll();
        public void OnFormatObjects(Office.IRibbonControl control) => FormatObjects?.Invoke();
        public void OnUndoBatch(Office.IRibbonControl control) => UndoBatch?.Invoke();
        public void OnVerifyResolved(Office.IRibbonControl control) => _verifyResolved();
        public void OnPreviousIssue(Office.IRibbonControl control) => _navigateIssue(-1);
        public void OnNextIssue(Office.IRibbonControl control) => _navigateIssue(1);
        public void OnPreviousOccurrence(Office.IRibbonControl control) => _navigateOccurrence(-1);
        public void OnNextOccurrence(Office.IRibbonControl control) => _navigateOccurrence(1);
        public void OnToggleIssues(Office.IRibbonControl control, bool pressed) => _setIssuesVisible(pressed);

        private static bool SafeGet(Func<bool> value)
        {
            try { return value(); }
            catch (Exception) { return false; }
        }
    }
}
