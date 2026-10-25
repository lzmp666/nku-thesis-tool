using System;

namespace ThesisAddIn.VSTO.Ui
{
    // No Word COM traffic when the pointer is idle. Hover runs once per dwell;
    // a native Resolve candidate is checked only briefly after pointer activity.
    internal sealed class DocumentInteractionGate
    {
        private int _x=int.MinValue, _y=int.MinValue;
        private bool _pressed, _pending;
        private DateTime _activity=DateTime.MinValue, _poll=DateTime.MinValue;
        internal bool Observe(int x,int y,bool pressed,DateTime now)
        {
            if (_x!=x || _y!=y || _pressed!=pressed) {
                _x=x; _y=y; _pressed=pressed; _activity=now; _pending=true;
                return false;
            }
            if (pressed || !_pending || (now-_activity).TotalMilliseconds<200) return false;
            _pending=false; return true;
        }
        internal bool CanCheckResolve(DateTime now)
        {
            if (_pressed || (now-_activity).TotalMilliseconds>1800 || (now-_poll).TotalMilliseconds<750) return false;
            _poll=now; return true;
        }
    }
}
