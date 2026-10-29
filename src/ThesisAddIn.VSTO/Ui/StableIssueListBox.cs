using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Windows.Forms;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>Buffered, pixel-scrolling cards without native ListBox selection painting.</summary>
    internal sealed class CardPointerEventArgs : EventArgs
    {
        internal int Index, X, Y;
        internal bool Handled;
    }
    internal sealed class CardHoverEventArgs : EventArgs
    {
        internal int Index;
    }
    internal sealed class StableIssueListBox : Control
    {
        private readonly IssueListViewport _viewport = new IssueListViewport();
        private readonly VScrollBar _scrollBar = new VScrollBar { Dock = DockStyle.Right, TabStop = false };
        private readonly Timer _scrollTimer = new Timer { Interval = 15 };
        private readonly Timer _hoverTimer = new Timer { Interval = 100 };
        private int _targetOffset;
        private int _updateDepth;
        private bool _settingScrollBar;
        private bool _measuring;
        private int _hoverCandidate = -1;
        private int _hoveredIndex = -1;

        internal StableIssueListBox()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            TabStop = true;
            Items = new CardItems(this);
            Controls.Add(_scrollBar);
            _scrollBar.ValueChanged += (sender, args) =>
            {
                if (_settingScrollBar) return;
                _scrollTimer.Stop();
                _targetOffset = _scrollBar.Value;
                SetOffset(_targetOffset);
            };
            _scrollTimer.Tick += (sender, args) =>
            {
                int distance = _targetOffset - _viewport.Offset;
                int step = Math.Max(1, (int)Math.Ceiling(Math.Abs(distance) * 0.4));
                SetOffset(Math.Abs(distance) <= 1 ? _targetOffset : _viewport.Offset + Math.Sign(distance) * step);
                if (_viewport.Offset == _targetOffset) _scrollTimer.Stop();
            };
            _hoverTimer.Tick += (sender, args) =>
            {
                _hoverTimer.Stop();
                if (_hoverCandidate >= 0 && _hoverCandidate != _hoveredIndex)
                {
                    _hoveredIndex = _hoverCandidate;
                    ItemHoverChanged?.Invoke(this, new CardHoverEventArgs { Index = _hoveredIndex });
                }
            };
        }
        public Collection<string> Items { get; }
        public int ContentWidth => Math.Max(1, ClientSize.Width - SystemInformation.VerticalScrollBarWidth);
        internal int ScrollOffset => _viewport.Offset;
        internal void RestoreScrollOffset(int offset) { _scrollTimer.Stop(); _targetOffset=offset; SetOffset(offset); }
        internal void RefreshLayout() => MeasureCards();
        public event EventHandler SelectedIndexChanged;
        public event EventHandler CurrentItemInvoked;
        internal event EventHandler<CardPointerEventArgs> ItemPointerPressed;
        internal event EventHandler<CardHoverEventArgs> ItemHoverChanged;
        public event DrawItemEventHandler DrawItem;
        public event MeasureItemEventHandler MeasureItem;
        public int SelectedIndex { get => _viewport.SelectedIndex; set => SelectCard(value, false); }
        internal void RevealSelection()
        {
            if (SelectedIndex >= 0) SelectCard(SelectedIndex, true);
        }
        public int TopIndex
        {
            get => _viewport.TopIndex;
            set
            {
                if (value < 0 || value >= Items.Count) return;
                _scrollTimer.Stop();
                _targetOffset = _viewport.ItemTop(value);
                SetOffset(_targetOffset);
            }
        }
        public void BeginUpdate() { _updateDepth++; }
        public void EndUpdate()
        {
            if (_updateDepth > 0) _updateDepth--;
            if (_updateDepth == 0) MeasureCards();
        }
        public Rectangle GetItemRectangle(int index) => index < 0 || index >= _viewport.Count
            ? Rectangle.Empty : new Rectangle(0, _viewport.ItemTop(index) - _viewport.Offset,
                ContentWidth, _viewport.ItemHeight(index));

        private void ItemsChanged()
        {
            _scrollTimer.Stop();
            _viewport.Select(-1);
            if (_updateDepth == 0) MeasureCards();
        }
        private void MeasureCards()
        {
            if (_measuring || _updateDepth > 0 || IsDisposed) return;
            _measuring = true;
            try
            {
                var heights = new List<int>();
                using (Graphics graphics = CreateGraphics())
                    for (int index = 0; index < Items.Count; index++)
                    {
                        var args = new MeasureItemEventArgs(graphics, index, 120);
                        MeasureItem?.Invoke(this, args);
                        heights.Add(Math.Max(1, args.ItemHeight));
                    }
                _viewport.SetLayout(heights, ClientSize.Height);
                _targetOffset = _viewport.Offset;
                UpdateScrollBar();
                Invalidate();
            }
            finally { _measuring = false; }
        }
        private void UpdateScrollBar()
        {
            _settingScrollBar = true;
            try
            {
                _scrollBar.Visible = _viewport.MaximumOffset > 0;
                _scrollBar.Minimum = 0;
                _scrollBar.Maximum = Math.Max(0, Math.Max(_viewport.TotalHeight, ClientSize.Height) - 1);
                _scrollBar.LargeChange = Math.Max(1, ClientSize.Height);
                _scrollBar.SmallChange = Math.Max(1, Font.Height * 2);
                _scrollBar.Value = _viewport.Offset;
            }
            finally { _settingScrollBar = false; }
        }
        private void SetOffset(int offset)
        {
            if (!_viewport.ScrollTo(offset)) return;
            UpdateScrollBar();
            Invalidate(new Rectangle(0, 0, ContentWidth, ClientSize.Height));
            QueueHoverAtLastPointer();
        }
        private void QueueHoverAtLastPointer()
        {
            if (!IsHandleCreated || IsDisposed) return;
            Point point = PointToClient(Cursor.Position);
            if (!ClientRectangle.Contains(point) || point.X >= ContentWidth) { ClearHover(); return; }
            QueueHover(point);
        }
        private void QueueHover(Point point)
        {
            int index = _viewport.HitTest(point.Y);
            if (index == _hoveredIndex && index == _hoverCandidate) return;
            _hoverCandidate = index;
            _hoverTimer.Stop();
            if (index != _hoveredIndex && index >= 0) _hoverTimer.Start();
            else if (_hoveredIndex >= 0) ClearHover();
        }
        private void ClearHover()
        {
            _hoverTimer.Stop();
            _hoverCandidate = -1;
            if (_hoveredIndex < 0) return;
            _hoveredIndex = -1;
            ItemHoverChanged?.Invoke(this, new CardHoverEventArgs { Index = -1 });
        }
        private void SelectCard(int index, bool ensureVisible)
        {
            int previous = _viewport.SelectedIndex;
            int oldOffset = _viewport.Offset;
            bool changed = _viewport.Select(index, ensureVisible);
            if (!changed && oldOffset == _viewport.Offset) return;
            if (oldOffset != _viewport.Offset)
            {
                _targetOffset = _viewport.Offset;
                UpdateScrollBar();
                Invalidate();
            }
            else
            {
                if (previous >= 0) Invalidate(GetItemRectangle(previous));
                if (index >= 0) Invalidate(GetItemRectangle(index));
            }
            if (changed)
            {
                AccessibilityNotifyClients(AccessibleEvents.Selection, index);
                SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
            }
        }
        protected override void OnPaintBackground(PaintEventArgs e) { /* Painted atomically in the buffered foreground. */ }
        protected override void OnPaint(PaintEventArgs e)
        {
            if (_updateDepth > 0) return;
            using (var background = new SolidBrush(BackColor)) e.Graphics.FillRectangle(background,e.ClipRectangle);
            var state = e.Graphics.Save();
            try
            {
                e.Graphics.SetClip(new Rectangle(0, 0, ContentWidth, ClientSize.Height),System.Drawing.Drawing2D.CombineMode.Intersect);
                for (int index = Math.Max(0, _viewport.TopIndex); index < _viewport.Count; index++)
                {
                    Rectangle bounds = GetItemRectangle(index);
                    if (bounds.Top >= ClientSize.Height) break;
                    if (bounds.Bottom <= 0 || !bounds.IntersectsWith(e.ClipRectangle)) continue;
                    DrawItemState flags = index == SelectedIndex ? DrawItemState.Selected : DrawItemState.None;
                    if (Focused && index == SelectedIndex) flags |= DrawItemState.Focus;
                    DrawItem?.Invoke(this, new DrawItemEventArgs(e.Graphics, Font, bounds, index, flags, ForeColor, BackColor));
                }
            }
            finally { e.Graphics.Restore(state); }
        }
        protected override void OnResize(EventArgs e) { base.OnResize(e); MeasureCards(); }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (e.X < ContentWidth) QueueHover(e.Location);
            else ClearHover();
        }
        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            ClearHover();
        }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            ClearHover();
            if (e.Button != MouseButtons.Left || e.X >= ContentWidth) return;
            _scrollTimer.Stop();
            _targetOffset = _viewport.Offset;
            int index = _viewport.HitTest(e.Y);
            var pointer = new CardPointerEventArgs { Index=index, X=e.X, Y=e.Y };
            ItemPointerPressed?.Invoke(this,pointer);
            if(pointer.Handled) return;
            Focus();
            if (index >= 0)
            {
                bool same = index == SelectedIndex;
                SelectCard(index, false);
                if (same) CurrentItemInvoked?.Invoke(this, EventArgs.Empty);
            }
        }
        protected override void OnMouseWheel(MouseEventArgs e)
        {
            ClearHover();
            if (e is HandledMouseEventArgs handled) handled.Handled = true;
            int wheelPixels = Math.Max(24, Font.Height * 3);
            _targetOffset = Math.Max(0, Math.Min(_viewport.MaximumOffset,
                _targetOffset - (int)Math.Round(e.Delta / 120.0 * wheelPixels)));
            if (_targetOffset != _viewport.Offset) _scrollTimer.Start();
        }
        protected override bool IsInputKey(Keys keyData)
        {
            switch (keyData & Keys.KeyCode)
            {
                case Keys.Up: case Keys.Down: case Keys.Home: case Keys.End:
                case Keys.PageUp: case Keys.PageDown: return true;
                default: return base.IsInputKey(keyData);
            }
        }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.Handled || Items.Count == 0) return;
            int next = SelectedIndex;
            switch (e.KeyCode)
            {
                case Keys.Up: next = Math.Max(0, next - 1); break;
                case Keys.Down: next = Math.Min(Items.Count - 1, next + 1); break;
                case Keys.Home: next = 0; break;
                case Keys.End: next = Items.Count - 1; break;
                case Keys.PageUp: next = Math.Max(0, _viewport.FindAt(Math.Max(0, _viewport.Offset - ClientSize.Height))); break;
                case Keys.PageDown: next = _viewport.FindAt(Math.Min(_viewport.TotalHeight - 1, _viewport.Offset + ClientSize.Height)); break;
                default: return;
            }
            _scrollTimer.Stop();
            SelectCard(next, true);
            e.Handled = true;
        }
        protected override AccessibleObject CreateAccessibilityInstance() => new CardListAccessibility(this);
        private sealed class CardListAccessibility : ControlAccessibleObject
        {
            private readonly StableIssueListBox _list;
            internal CardListAccessibility(StableIssueListBox list) : base(list) { _list = list; }
            public override AccessibleRole Role => AccessibleRole.List;
            public override int GetChildCount() => _list.Items.Count;
            public override AccessibleObject GetChild(int index) => index >= 0 && index < _list.Items.Count
                ? new CardAccessibility(_list, index) : null;
            public override AccessibleObject GetSelected() => GetChild(_list.SelectedIndex);
        }
        private sealed class CardAccessibility : AccessibleObject
        {
            private readonly StableIssueListBox _list;
            private readonly int _index;
            internal CardAccessibility(StableIssueListBox list, int index) { _list = list; _index = index; }
            public override string Name { get => _list.Items[_index]; set { } }
            public override AccessibleRole Role => AccessibleRole.ListItem;
            public override AccessibleObject Parent => _list.AccessibilityObject;
            public override Rectangle Bounds => _list.RectangleToScreen(_list.GetItemRectangle(_index));
            public override AccessibleStates State => AccessibleStates.Selectable | AccessibleStates.Focusable
                | (_list.SelectedIndex == _index ? AccessibleStates.Selected : AccessibleStates.None)
                | (_list.ClientRectangle.IntersectsWith(_list.GetItemRectangle(_index)) ? AccessibleStates.None : AccessibleStates.Offscreen);
            public override string DefaultAction => "定位问题";
            public override void DoDefaultAction() { _list.Focus(); _list.SelectCard(_index, true); }
        }
        private sealed class CardItems : Collection<string>
        {
            private readonly StableIssueListBox _owner;
            internal CardItems(StableIssueListBox owner) { _owner = owner; }
            protected override void InsertItem(int index, string item) { base.InsertItem(index, item); _owner.ItemsChanged(); }
            protected override void ClearItems() { base.ClearItems(); _owner.ItemsChanged(); }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { _scrollTimer.Stop(); _scrollTimer.Dispose(); _hoverTimer.Stop(); _hoverTimer.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
