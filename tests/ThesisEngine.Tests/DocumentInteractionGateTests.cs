using ThesisAddIn.VSTO.Ui;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DocumentInteractionGateTests
{
    [Fact]
    public void HoverRunsOnceAfterDwellAndNeverRepeatsForAnIdlePointer()
    {
        var gate=new DocumentInteractionGate(); var now=new DateTime(2026,10,25);
        Assert.False(gate.Observe(10,20,false,now));
        Assert.False(gate.Observe(10,20,false,now.AddMilliseconds(100)));
        Assert.True(gate.Observe(10,20,false,now.AddMilliseconds(250)));
        for(int i=2;i<100;i++) Assert.False(gate.Observe(10,20,false,now.AddMilliseconds(250*i)));
        Assert.False(gate.CanCheckResolve(now.AddSeconds(30)));
    }
    [Fact]
    public void MotionAndPressedMouseDoNotStartWordHoverQueries()
    {
        var gate=new DocumentInteractionGate(); var now=new DateTime(2026,10,25);
        for(int i=0;i<20;i++) Assert.False(gate.Observe(i,20,false,now.AddMilliseconds(250*i)));
        Assert.False(gate.Observe(20,20,true,now.AddSeconds(5)));
        Assert.False(gate.Observe(20,20,true,now.AddSeconds(6)));
        Assert.False(gate.CanCheckResolve(now.AddSeconds(6)));
        Assert.False(gate.Observe(20,20,false,now.AddSeconds(7)));
        Assert.True(gate.Observe(20,20,false,now.AddMilliseconds(7250)));
    }
    [Fact]
    public void ResolveChecksAreBriefAndThrottledAfterPointerActivity()
    {
        var gate=new DocumentInteractionGate(); var now=new DateTime(2026,10,25);
        gate.Observe(10,20,false,now);
        Assert.True(gate.CanCheckResolve(now.AddMilliseconds(250)));
        Assert.False(gate.CanCheckResolve(now.AddMilliseconds(500)));
        Assert.True(gate.CanCheckResolve(now.AddMilliseconds(1000)));
        Assert.False(gate.CanCheckResolve(now.AddSeconds(2)));
    }
}
