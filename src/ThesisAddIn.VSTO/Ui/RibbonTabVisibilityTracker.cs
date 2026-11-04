using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Accessibility;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>Observe native Ribbon tab selection without scanning Word's document UIA tree.</summary>
    internal sealed class RibbonTabVisibilityTracker : IDisposable
    {
        private const string TabName = "论文格式检查";
        private readonly Dictionary<int, Action<bool>> _callbacks = new Dictionary<int, Action<bool>>();
        private readonly Dictionary<int, bool> _lastSelection = new Dictionary<int, bool>();
        private readonly WinEventCallback _callback;
        private IntPtr _hook;

        internal RibbonTabVisibilityTracker()
        {
            _callback = OnWinEvent;
            // Out-of-context callbacks run on the registering Word UI thread.
            _hook = SetWinEventHook(0x8005, 0x800A, IntPtr.Zero, _callback,
                (uint)Process.GetCurrentProcess().Id, 0, 0);
        }

        internal void Track(int windowHandle, Action<bool> selectedChanged)
        {
            if (selectedChanged != null) _callbacks[windowHandle] = selectedChanged;
        }

        private void OnWinEvent(IntPtr hook, uint eventId, IntPtr window, int objectId,
            int childId, uint thread, uint timestamp)
        {
            if (window == IntPtr.Zero || _hook == IntPtr.Zero) return;
            try
            {
                int frame = GetAncestor(window, 2).ToInt32();
                if (!_callbacks.TryGetValue(frame, out Action<bool> selectedChanged)) return;
                var className = new StringBuilder(64);
                GetClassName(window, className, className.Capacity);
                // Ignore document text and our WinForms list before any accessibility call.
                if (className.ToString() != "NetUIHWND" && className.ToString() != "OpusApp") return;
                if (AccessibleObjectFromEvent(window, objectId, childId,
                    out IAccessible accessible, out object child) < 0 || accessible == null) return;
                if (Convert.ToInt32(accessible.get_accRole(child)) != 0x25) return; // ROLE_SYSTEM_PAGETAB
                if ((Convert.ToInt32(accessible.get_accState(child)) & 0x2) == 0) return;
                string name = accessible.get_accName(child);
                if (name == TabName)
                {
                    NotifySelection(frame, true, selectedChanged);
                    return;
                }

                // Built-in Word tabs are exposed as separate COM children on some
                // Office builds; get_accName(index) is not supported by that provider.
                if (IsBuiltInRibbonTab(name))
                {
                    NotifySelection(frame, false, selectedChanged);
                    return;
                }

                // Only sibling tabs of our custom Ribbon tab may hide the pane.
                IAccessible siblings = child is int index && index > 0
                    ? accessible : accessible.accParent as IAccessible;
                if (siblings == null) return;
                for (int i = 1; i <= Math.Min(64, siblings.accChildCount); i++)
                {
                    string siblingName = null;
                    try
                    {
                        var sibling = siblings.get_accChild(i) as IAccessible;
                        siblingName = sibling == null ? siblings.get_accName(i) : sibling.get_accName(0);
                    }
                    catch (COMException) { continue; }
                    if (siblingName == TabName)
                    {
                        NotifySelection(frame, false, selectedChanged);
                        return;
                    }
                }
            }
            catch (Exception) { /* Office may destroy an accessible element during a tab transition. */ }
        }

        private void NotifySelection(int frame, bool selected, Action<bool> selectedChanged)
        {
            if (_lastSelection.TryGetValue(frame, out bool previous) && previous == selected) return;
            _lastSelection[frame] = selected;
            selectedChanged(selected);
        }

        private static bool IsBuiltInRibbonTab(string name)
        {
            switch (name)
            {
                case "开始": case "插入": case "设计": case "布局": case "页面布局":
                case "引用": case "邮件": case "审阅": case "视图": case "帮助":
                case "开发工具": case "绘图": case "Home": case "Insert": case "Design":
                case "Layout": case "Page Layout": case "References": case "Mailings":
                case "Review": case "View": case "Help": case "Developer": case "Draw":
                    return true;
                default: return false;
            }
        }

        public void Dispose()
        {
            if (_hook != IntPtr.Zero) UnhookWinEvent(_hook);
            _hook = IntPtr.Zero;
            _callbacks.Clear();
            _lastSelection.Clear();
        }

        private delegate void WinEventCallback(IntPtr hook, uint eventId, IntPtr window,
            int objectId, int childId, uint thread, uint timestamp);
        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr module,
            WinEventCallback callback, uint processId, uint threadId, uint flags);
        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hook);
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int maximum);
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromEvent(IntPtr window, int objectId, int childId,
            [MarshalAs(UnmanagedType.Interface)] out IAccessible accessible,
            [MarshalAs(UnmanagedType.Struct)] out object child);
    }
}
