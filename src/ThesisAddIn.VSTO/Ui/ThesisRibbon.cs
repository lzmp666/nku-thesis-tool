using System;
using System.Runtime.InteropServices;
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
        private readonly Action _showIssues;
        private readonly Func<bool> _canFixCurrent;
        private readonly Func<bool> _canFixAll;
        private readonly Func<bool> _canVerifyResolved;
        private readonly Func<bool> _canNavigateIssue;
        private readonly Func<bool> _canNavigateOccurrence;
        private Office.IRibbonUI _ribbonUi;

        public ThesisRibbon(Action analyze, Action fixCurrent, Action fixAll, Action verifyResolved,
            Action<int> navigateIssue, Action<int> navigateOccurrence, Action showIssues,
            Func<bool> canFixCurrent, Func<bool> canFixAll, Func<bool> canVerifyResolved,
            Func<bool> canNavigateIssue, Func<bool> canNavigateOccurrence)
        {
            _analyze = analyze ?? throw new ArgumentNullException(nameof(analyze));
            _fixCurrent = fixCurrent ?? throw new ArgumentNullException(nameof(fixCurrent));
            _fixAll = fixAll ?? throw new ArgumentNullException(nameof(fixAll));
            _verifyResolved = verifyResolved ?? throw new ArgumentNullException(nameof(verifyResolved));
            _navigateIssue = navigateIssue ?? throw new ArgumentNullException(nameof(navigateIssue));
            _navigateOccurrence = navigateOccurrence ?? throw new ArgumentNullException(nameof(navigateOccurrence));
            _showIssues = showIssues ?? throw new ArgumentNullException(nameof(showIssues));
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
        <group id=""NkuThesisCheck"" label=""检查"">
          <button id=""NkuAnalyzeCurrent"" label=""分析当前文档"" size=""large""
                  screentip=""分析当前文档""
                  supertip=""分析当前 Word 文档，并在本页更新格式问题标注。""
                  onAction=""OnAnalyzeCurrent"" />
          <button id=""NkuShowIssues"" label=""问题检索""
                  screentip=""显示问题检索列表""
                  supertip=""打开右侧问题检索列表；不创建新文档。""
                  onAction=""OnShowIssues"" />
        </group>
        <group id=""NkuThesisRepair"" label=""修复"">
          <button id=""NkuFixCurrent"" label=""修复当前问题"" getEnabled=""GetEnabled""
                  screentip=""预览并修复当前问题""
                  supertip=""先查看将修改的格式和值，确认后更新当前文档。""
                  onAction=""OnFixCurrent"" />
          <button id=""NkuFixAll"" label=""预览并修复全部"" getEnabled=""GetEnabled""
                  screentip=""预览并修复全部问题""
                  supertip=""预览所有可修复的问题，确认后一次性更新当前文档。""
                  onAction=""OnFixAll"" />
          <button id=""NkuVerifyResolved"" label=""验证已解决"" getEnabled=""GetEnabled""
                  screentip=""重新检查当前问题位置""
                  supertip=""你手动调整后重新分析此位置；确认符合要求才清除标记，仍有问题则保留。""
                  onAction=""OnVerifyResolved"" />
        </group>
        <group id=""NkuThesisNavigation"" label=""定位"">
          <button id=""NkuPreviousIssue"" label=""上一项"" getEnabled=""GetEnabled"" onAction=""OnPreviousIssue"" />
          <button id=""NkuNextIssue"" label=""下一项"" getEnabled=""GetEnabled"" onAction=""OnNextIssue"" />
          <button id=""NkuPreviousOccurrence"" label=""上一处"" getEnabled=""GetEnabled"" onAction=""OnPreviousOccurrence"" />
          <button id=""NkuNextOccurrence"" label=""下一处"" getEnabled=""GetEnabled"" onAction=""OnNextOccurrence"" />
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

        public void OnAnalyzeCurrent(Office.IRibbonControl control) => _analyze();
        public void OnFixCurrent(Office.IRibbonControl control) => _fixCurrent();
        public void OnFixAll(Office.IRibbonControl control) => _fixAll();
        public void OnVerifyResolved(Office.IRibbonControl control) => _verifyResolved();
        public void OnPreviousIssue(Office.IRibbonControl control) => _navigateIssue(-1);
        public void OnNextIssue(Office.IRibbonControl control) => _navigateIssue(1);
        public void OnPreviousOccurrence(Office.IRibbonControl control) => _navigateOccurrence(-1);
        public void OnNextOccurrence(Office.IRibbonControl control) => _navigateOccurrence(1);
        public void OnShowIssues(Office.IRibbonControl control) => _showIssues();

        private static bool SafeGet(Func<bool> value)
        {
            try { return value(); }
            catch (Exception) { return false; }
        }
    }
}
