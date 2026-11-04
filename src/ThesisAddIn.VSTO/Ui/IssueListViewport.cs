using System;
using System.Collections.Generic;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>Pixel layout: selecting a card never snaps the viewport.</summary>
    internal sealed class IssueListViewport
    {
        private readonly List<int> _tops = new List<int>();
        private readonly List<int> _heights = new List<int>();
        internal int Offset { get; private set; }
        internal int Height { get; private set; }
        internal int TotalHeight { get; private set; }
        internal int SelectedIndex { get; private set; } = -1;
        internal int Count => _heights.Count;
        internal int MaximumOffset => Math.Max(0, TotalHeight - Height);
        internal int TopIndex => FindAt(Offset);

        internal void SetLayout(IEnumerable<int> heights, int height)
        {
            _tops.Clear(); _heights.Clear(); TotalHeight = 0;
            Height = Math.Max(0, height);
            foreach (int itemHeight in heights)
            {
                _tops.Add(TotalHeight);
                int normalized = Math.Max(1, itemHeight);
                _heights.Add(normalized);
                TotalHeight = checked(TotalHeight + normalized);
            }
            if (SelectedIndex >= Count) SelectedIndex = -1;
            ScrollTo(Offset);
        }
        internal bool ScrollTo(int offset)
        {
            int next = Math.Max(0, Math.Min(MaximumOffset, offset));
            if (next == Offset) return false;
            Offset = next; return true;
        }
        internal bool Select(int index, bool ensureVisible = false)
        {
            if (index < -1 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
            bool changed = index != SelectedIndex;
            SelectedIndex = index;
            if (ensureVisible && index >= 0)
            {
                if (_tops[index] < Offset) ScrollTo(_tops[index]);
                else if (_tops[index] + _heights[index] > Offset + Height)
                    ScrollTo(_heights[index] > Height ? _tops[index] : _tops[index] + _heights[index] - Height);
            }
            return changed;
        }
        internal int ItemTop(int index) => _tops[index];
        internal int ItemHeight(int index) => _heights[index];
        internal int HitTest(int y) => y < 0 || y >= Height ? -1 : FindAt(Offset + y);
        internal int FindAt(int offset)
        {
            if (offset < 0 || offset >= TotalHeight) return -1;
            int low = 0, high = Count - 1;
            while (low <= high)
            {
                int middle = low + (high - low) / 2;
                if (offset < _tops[middle]) high = middle - 1;
                else if (offset >= _tops[middle] + _heights[middle]) low = middle + 1;
                else return middle;
            }
            return -1;
        }
    }
}
