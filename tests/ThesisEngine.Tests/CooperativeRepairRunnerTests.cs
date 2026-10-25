using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class CooperativeRepairRunnerTests
{
    [Fact]
    public async Task UserEditDuringYieldStopsBeforeTheNextWrite()
    {
        bool edited = false;
        var writes = new List<int>();
        await Assert.ThrowsAsync<InvalidOperationException>(() => CooperativeRepairRunner.RunAsync(
            Enumerable.Range(0, 20).ToArray(),
            () => { if (edited) throw new InvalidOperationException("edited"); },
            writes.Add, () => { edited = true; return Task.CompletedTask; }));
        Assert.InRange(writes.Count, 1, 4);
        Assert.Equal(Enumerable.Range(0, writes.Count), writes);
    }

    [Fact]
    public async Task LargeBatchYieldsAndValidatesBetweenBoundedSlices()
    {
        var writes = new List<int>(); int yields = 0, validations = 0;
        await CooperativeRepairRunner.RunAsync(Enumerable.Range(0, 41).ToArray(),
            () => validations++, writes.Add, () => { yields++; return Task.CompletedTask; });
        Assert.True(yields >= 10);
        Assert.Equal(yields + 1, validations);
        Assert.Equal(Enumerable.Range(0, 41), writes);
    }

    [Fact]
    public async Task FailedWriteIsNotRetriedAndNoLaterWriteRuns()
    {
        var writes = new List<int>();
        await Assert.ThrowsAsync<IOException>(() => CooperativeRepairRunner.RunAsync(
            Enumerable.Range(0, 30).ToArray(), () => { }, value => {
                if (value == 5) throw new IOException("word failure");
                writes.Add(value);
            }, () => Task.CompletedTask));
        Assert.Equal(Enumerable.Range(0, 5), writes);
    }
}
