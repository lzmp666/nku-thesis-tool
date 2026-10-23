using System.Windows.Forms;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>Mouse selection must not snap a variable-height card to the top.</summary>
    internal sealed class StableIssueListBox : ListBox
    {
        protected override void WndProc(ref Message message)
        {
            const int LeftButtonDown = 0x0201;
            if (message.Msg != LeftButtonDown || Items.Count == 0)
            {
                base.WndProc(ref message);
                return;
            }

            int firstVisible = TopIndex;
            base.WndProc(ref message);
            if (!IsDisposed && IsHandleCreated && firstVisible < Items.Count && TopIndex != firstVisible)
                TopIndex = firstVisible;
        }
    }
}
