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
        <group id=""NkuThesisCheck"" label=""检查与检索"">
          <button id=""NkuAnalyzeCurrent"" label=""分析当前文档"" size=""large"" getImage=""GetButtonImage""
                  screentip=""分析当前文档""
                  supertip=""分析当前 Word 文档，并在本页更新格式问题标注。""
                  onAction=""OnAnalyzeCurrent"" />
          <toggleButton id=""NkuShowIssues"" getLabel=""GetIssuesLabel"" size=""large""
                        getImage=""GetButtonImage"" getPressed=""GetIssuesPressed""
                        screentip=""打开或关闭问题检索列表""
                        supertip=""在当前 Word 窗口显示或隐藏右侧问题检索列表。""
                        onAction=""OnToggleIssues"" />
        </group>
        <group id=""NkuThesisRepair"" label=""修复"">
          <button id=""NkuFixCurrent"" label=""修复当前问题"" size=""large"" getImage=""GetButtonImage"" getEnabled=""GetEnabled""
                  screentip=""预览并修复当前问题""
                  supertip=""先查看将修改的格式和值，确认后更新当前文档。""
                  onAction=""OnFixCurrent"" />
          <button id=""NkuFixAll"" label=""批量修复预览"" size=""large"" getImage=""GetButtonImage"" getEnabled=""GetEnabled""
                  screentip=""预览并修复全部问题""
                  supertip=""预览所有可修复的问题，确认后一次性更新当前文档。""
                  onAction=""OnFixAll"" />
          <button id=""NkuVerifyResolved"" label=""验证已解决"" getImage=""GetButtonImage"" getEnabled=""GetEnabled""
                  screentip=""重新检查当前问题位置""
                  supertip=""你手动调整后重新分析此位置；确认符合要求才清除标记，仍有问题则保留。""
                  onAction=""OnVerifyResolved"" />
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
            try { _ribbonUi?.Invalidate(); }
            catch (COMException) { }
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "NkuFixCurrent": return SafeGet(_canFixCurrent);
                case "NkuFixAll": return SafeGet(_canFixAll);
                case "NkuVerifyResolved": return SafeGet(_canVerifyResolved);
                case "NkuPreviousIssue":
                case "NkuNextIssue": return SafeGet(_canNavigateIssue);
                case "NkuPreviousOccurrence":
                case "NkuNextOccurrence": return SafeGet(_canNavigateOccurrence);
                default: return true;
            }
        }

        public bool GetIssuesPressed(Office.IRibbonControl control) => SafeGet(_isIssuesVisible);
        public string GetIssuesLabel(Office.IRibbonControl control) =>
            SafeGet(_isIssuesVisible) ? "关闭检索" : "打开检索";

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
                case "NkuVerifyResolved": index = 4; break;
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

        private sealed class RibbonPictureConverter : AxHost
        {
            private RibbonPictureConverter() : base(string.Empty) { }
            internal static stdole.IPictureDisp Convert(Image image) =>
                (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }

        public void OnAnalyzeCurrent(Office.IRibbonControl control) => _analyze();
        public void OnFixCurrent(Office.IRibbonControl control) => _fixCurrent();
        public void OnFixAll(Office.IRibbonControl control) => _fixAll();
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
