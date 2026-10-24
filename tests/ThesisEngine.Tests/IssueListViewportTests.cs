using ThesisAddIn.VSTO.Ui;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class IssueListViewportTests
{
    [Fact]
    public void MouseSelectionPreservesPartialCardOffsetAndDoesNotRepeatAnEvent()
    {
        var viewport = new IssueListViewport();
        viewport.SetLayout(new[] { 260, 310, 200 }, 400);
        viewport.ScrollTo(123);
        Assert.True(viewport.Select(1));
        Assert.Equal(123, viewport.Offset);
        Assert.False(viewport.Select(1));
        Assert.Equal(123, viewport.Offset);
        Assert.Equal(0, viewport.TopIndex);
    }
    [Fact]
    public void WheelAndHitTestingUsePixelsInsteadOfWholeCards()
    {
        var viewport = new IssueListViewport();
        viewport.SetLayout(new[] { 260, 310, 200 }, 400);
        viewport.ScrollTo(48);
        Assert.Equal(48, viewport.Offset);
        Assert.Equal(0, viewport.HitTest(211));
        Assert.Equal(1, viewport.HitTest(212));
        Assert.Equal(-1, viewport.HitTest(400));
        Assert.Equal(-1, viewport.HitTest(-1));
    }
    [Fact]
    public void ResizingClampsOffsetAndKeepsAValidSelection()
    {
        var viewport = new IssueListViewport();
        viewport.SetLayout(new[] { 260, 310, 200 }, 300);
        viewport.ScrollTo(470);
        viewport.Select(2);
        viewport.SetLayout(new[] { 260, 310, 200 }, 600);
        Assert.Equal(170, viewport.Offset);
        Assert.Equal(2, viewport.SelectedIndex);
        viewport.SetLayout(new[] { 260 }, 600);
        Assert.Equal(0, viewport.Offset);
        Assert.Equal(-1, viewport.SelectedIndex);
    }
    [Fact]
    public void OnlyExplicitKeyboardNavigationEnsuresVisibility()
    {
        var viewport = new IssueListViewport();
        viewport.SetLayout(new[] { 260, 310, 200 }, 400);
        viewport.Select(2);
        Assert.Equal(0, viewport.Offset);
        viewport.Select(2, ensureVisible: true);
        Assert.Equal(370, viewport.Offset);
        viewport.Select(0, ensureVisible: true);
        Assert.Equal(0, viewport.Offset);
    }
}
