using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>为已关闭的任务窗格提供始终可找到的 Word 功能区入口。</summary>
    [ComVisible(true)]
    public sealed class ThesisRibbon : Office.IRibbonExtensibility
    {
        private readonly Action _openPane;

        public ThesisRibbon(Action openPane)
        {
            _openPane = openPane ?? throw new ArgumentNullException(nameof(openPane));
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
  <ribbon>
    <tabs>
      <tab id=""NkuThesisTab"" label=""论文格式检查"" insertAfterMso=""TabHome"">
        <group id=""NkuThesisCommands"" label=""NKU Thesis Tool"">
          <button id=""NkuOpenDiagnostics"" label=""打开检查面板""
                  screentip=""打开论文格式检查""
                  supertip=""显示当前文档的检查、定位和修复操作。关闭侧栏后，可从这里重新打开。""
                  onAction=""OnOpenPane"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnOpenPane(Office.IRibbonControl control)
        {
            _openPane();
        }
    }
}
