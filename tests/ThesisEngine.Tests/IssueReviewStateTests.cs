using ThesisAddIn.Protocol;
using ThesisAddIn.VSTO.Ui;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class IssueReviewStateTests
{
    [Fact]
    public void UndoBatchSelectsOnlyCompletedRowsAndReopensOnlyTheSelection()
    {
        var state=new IssueReviewState(); var a=Issue("font","a"); var b=Issue("size","b"); var c=Issue("caption","c");
        state.Refresh(new[]{a,b,c}); state.Complete(new[]{a,b});
        state.SetBatchMode(true,true); state.SelectAll(true);
        Assert.Equal(2,state.Checked.Count);
        state.Toggle(c); Assert.Equal(2,state.Checked.Count);
        state.Toggle(b); Assert.Single(state.Checked);
        state.ReopenMany(state.Checked.ToArray());
        Assert.False(state.IsRepaired(a)); Assert.True(state.IsRepaired(b));
        Assert.Empty(state.Checked);
    }
    [Fact]
    public void BatchModeSwitchClearsIneligibleChecksAndRefreshDropsExternallyUndoneRows()
    {
        var state=new IssueReviewState(); var a=Issue("font","a"); var b=Issue("size","b");
        state.Refresh(new[]{a,b}); state.Complete(new[]{a});
        state.SetBatch(true); state.SelectAll(true); Assert.Contains(IssueReviewState.Key(b),state.Checked);
        state.SetBatchMode(true,true); Assert.Empty(state.Checked);
        state.SelectAll(true); Assert.Contains(IssueReviewState.Key(a),state.Checked);
        state.Refresh(new[]{a,b}); Assert.Empty(state.Checked);
        state.SetBatch(false); Assert.False(state.BatchUndoMode);
    }
    private static EngineDiagnostic Issue(string rule, string location, string actual="Arial", string level="auto") =>
        new() { RuleId=rule, Location=location, Expected="Times New Roman", Actual=actual, FixLevel=level };
    [Fact]
    public void BatchChecksAreSharedAndExcludeCompletedOrManualIssues()
    {
        var state=new IssueReviewState();
        state.Refresh(new[]{Issue("font","a"),Issue("size","b"),Issue("manual","c",level:"manual")});
        state.SetBatch(true); state.SelectAll(true);
        Assert.Equal(2,state.Checked.Count);
        state.Complete(new[]{state.Rows[0]});
        Assert.Single(state.Checked);
        state.Toggle(state.Rows[0]); state.Toggle(state.Rows[2]);
        Assert.Single(state.Checked);
    }
    [Fact]
    public void CompletedIssuesStayInOriginalOrderForIndependentUndo()
    {
        var state=new IssueReviewState(); var a=Issue("font","a"); var b=Issue("size","b");
        state.Refresh(new[]{a,b}); state.Complete(new[]{a}); state.Refresh(new[]{b});
        Assert.Equal(new[]{"font","size"},state.Rows.Select(x=>x.RuleId));
        Assert.True(state.IsRepaired(state.Rows[0]));
        state.Reopen(IssueReviewState.Key(a)); state.Refresh(new[]{a,b});
        Assert.Empty(state.Repaired);
        Assert.Equal(2,state.Rows.Count);
    }
    [Fact]
    public void SameTargetIsOneIssueWithEveryOccurrenceAndActualValue()
    {
        var state=new IssueReviewState(); state.Refresh(new[]{Issue("font","a"),Issue("font","b","Calibri")});
        Assert.Single(state.Rows); Assert.Equal(2,state.Rows[0].Occurrences.Count);
        Assert.Contains("Arial",state.Rows[0].Actual); Assert.Contains("Calibri",state.Rows[0].Actual);
    }
    [Fact]
    public void ExternalUndoReopensAnIssueAndSelectionDoesNotRebuildData()
    {
        var state=new IssueReviewState(); var issue=Issue("font","a");
        state.Refresh(new[]{issue}); state.Complete(new[]{issue}); state.Refresh(new[]{issue});
        Assert.Empty(state.Repaired);
        int data=0,selection=0; state.Changed+=(s,e)=>data++; state.SelectionChanged+=(s,e)=>selection++;
        state.Select(issue,"a"); state.Select(issue,"a");
        Assert.Equal(0,data); Assert.Equal(1,selection);
    }
}
