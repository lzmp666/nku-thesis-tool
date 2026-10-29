param([Parameter(Mandatory=$true)][string]$AddinOutputDirectory)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System;
using System.Reflection;
public sealed class HoverEventCapture
{
    public int Calls;
    public string RuleId;
    public string Location;
    public void OnHover(object sender, EventArgs args)
    {
        Calls++;
        var flags = BindingFlags.Instance | BindingFlags.NonPublic;
        object diagnostic = args.GetType().GetField("Diagnostic", flags).GetValue(args);
        RuleId = diagnostic == null ? null : (string)diagnostic.GetType().GetProperty("RuleId").GetValue(diagnostic);
        Location = (string)args.GetType().GetField("Location", flags).GetValue(args);
    }
}
'@
[Reflection.Assembly]::LoadFrom((Join-Path $AddinOutputDirectory 'ThesisAddIn.Protocol.dll')) | Out-Null
$assembly=[Reflection.Assembly]::LoadFrom((Join-Path $AddinOutputDirectory 'ThesisAddIn.VSTO.dll'))
$flags=[Reflection.BindingFlags]'Instance,Public,NonPublic'
$search=New-Object ThesisAddIn.VSTO.Ui.DiagnosticsPane
$state=$search.GetType().GetField('Review',$flags).GetValue($search)
$suggestion=[Activator]::CreateInstance($search.GetType(),$flags,$null,@($true,$state),$null)
try {
    foreach($pane in @($search,$suggestion)) { $pane.Size=[Drawing.Size]::new(360,620); $pane.CreateControl() }
    $analysis=New-Object ThesisAddIn.Protocol.EngineAnalysisResult
    foreach($rule in @('demo/east-asian-font','demo/font-size','demo/western-font')) {
        $issue=New-Object ThesisAddIn.Protocol.EngineDiagnostic
        $issue.RuleId=$rule; $issue.Expected='12pt'; $issue.Actual='10.5pt'; $issue.FixLevel='auto'
        $issue.Location='word/document.xml#paragraph['+$analysis.Diagnostics.Count+']/run[0]'
        $analysis.Diagnostics.Add($issue)
    }
    $search.ShowAnalysis($analysis)
    $getList={ param($pane) $pane.GetType().GetField('_issueList',$flags).GetValue($pane) }
    $left=& $getList $search; $right=& $getList $suggestion
    if($left.Items.Count -ne 3 -or $right.Items.Count -ne 3) { throw 'Review regions do not share all issues.' }
    $hoverEvent=$suggestion.GetType().GetEvent('DiagnosticHoverChanged',$flags)
    $hoverCapture=[HoverEventCapture]::new()
    $hoverHandler=[Delegate]::CreateDelegate($hoverEvent.EventHandlerType,$hoverCapture,$hoverCapture.GetType().GetMethod('OnHover'))
    $hoverEvent.GetAddMethod($true).Invoke($suggestion,@($hoverHandler)) | Out-Null
    $move=$right.GetType().GetMethod('OnMouseMove',$flags)
    $firstBounds=$right.GetItemRectangle(0)
    $move.Invoke($right,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::None,0,20,$firstBounds.Top+12,0))) | Out-Null
    $hoverDeadline=[DateTime]::UtcNow.AddMilliseconds(350)
    while([DateTime]::UtcNow -lt $hoverDeadline) { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 10 }
    if($hoverCapture.Calls -lt 1 -or $hoverCapture.RuleId -ne 'demo/east-asian-font' -or
        $hoverCapture.Location -ne 'word/document.xml#paragraph[0]/run[0]') { throw 'Hover did not publish the corresponding suggestion and source location.' }
    $right.GetType().GetMethod('OnMouseLeave',$flags).Invoke($right,@([EventArgs]::Empty)) | Out-Null
    if($null -ne $hoverCapture.RuleId) { throw 'Leaving the suggestion card did not clear its hover preview.' }
    $state.GetType().GetMethod('SetBatch',$flags).Invoke($state,@($true)) | Out-Null
    $checked=$state.GetType().GetField('Checked',$flags).GetValue($state)
    $click=$left.GetType().GetMethod('OnMouseDown',$flags)
    $click.Invoke($left,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,14,18,0))) | Out-Null
    if($checked.Count -ne 1) { throw 'Search checkbox did not toggle the shared selection.' }
    $click.Invoke($right,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,14,18,0))) | Out-Null
    if($checked.Count -ne 0) { throw 'Suggestion checkbox did not update the same selection.' }
    $state.GetType().GetMethod('SetBatch',$flags).Invoke($state,@($false)) | Out-Null
    $script:repairs=0; $script:undos=0
    $search.add_RepairRequested({$script:repairs++})
    $search.add_UndoRequested({$script:undos++})
    $suggestion.add_RepairRequested({$script:repairs++})
    $suggestion.add_UndoRequested({$script:undos++})
    $bounds=$left.GetItemRectangle(0)
    $click.Invoke($left,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,30,$bounds.Bottom-18,0))) | Out-Null
    if($script:repairs -ne 0) { throw 'The searchable index should not duplicate source-card actions.' }
    $rightBounds=$right.GetItemRectangle(0)
    $click.Invoke($right,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,30,$rightBounds.Bottom-18,0))) | Out-Null
    if($script:repairs -ne 1) { throw 'Solve is not connected to the repair event.' }
    $selected=$search.SelectedDiagnostic
    $key=$assembly.GetType('ThesisAddIn.VSTO.Ui.IssueReviewState').GetMethod('Key',[Reflection.BindingFlags]'Static,NonPublic').Invoke($null,@($selected))
    $state.GetType().GetField('Repaired',$flags).GetValue($state).Add($key) | Out-Null
    $click.Invoke($right,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,100,$rightBounds.Bottom-18,0))) | Out-Null
    if($script:undos -ne 1) { throw 'Undo is not connected to the item undo event.' }
    $search.SetBusy($true,'checking')
    $suggestion.SetBusy($true,'checking')
    $query=$search.GetType().GetField('_searchBox',$flags).GetValue($search)
    if(-not $left.Enabled -or -not $query.Enabled) { throw 'Read-only browsing/search is blocked while repairing.' }
    $click.Invoke($right,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,100,$rightBounds.Bottom-18,0))) | Out-Null
    if($script:repairs -ne 1 -or $script:undos -ne 1) { throw 'A mutation was dispatched while a repair was already running.' }
    $search.GetType().GetMethod('OnSearchGotFocus',$flags).Invoke($search,@($query,[EventArgs]::Empty)) | Out-Null
    $query.Text='Times'
    if($left.Items.Count -ne 0) { throw 'Busy search did not filter immediately.' }
    $query.Text=''
    if($left.Items.Count -ne 3) { throw 'Busy search did not restore all rows.' }
    $search.SetBusy($false,$null)
    $suggestion.SetBusy($false,$null)
    # An equivalent fresh analysis must retain both search text and the exact
    # pixel scroll offset, with no transient clear/selection reset.
    $search.Height=210; $search.PerformLayout()
    $view=$left.GetType().GetField('_viewport',$flags).GetValue($left)
    $left.GetType().GetMethod('RestoreScrollOffset',$flags).Invoke($left,@(17)) | Out-Null
    $search.ShowAnalysis($analysis)
    if($view.GetType().GetProperty('Offset',$flags).GetValue($view) -ne 17) { throw 'Equivalent analysis moved the list viewport.' }
    $fontQuery=[string]::Concat([char]0x5B57,[char]0x4F53)
    $query.Text=$fontQuery
    $search.ShowAnalysis($analysis)
    if($query.Text -ne $fontQuery -or $left.Items.Count -ne 2) { throw 'Re-analysis reset the active search.' }
    $query.Text=''
    $completeArgs=New-Object object[] 1
    $completeArgs[0]=[ThesisAddIn.Protocol.EngineDiagnostic[]]@($selected)
    $state.GetType().GetMethod('Complete',$flags).Invoke($state,$completeArgs) | Out-Null
    $search.GetType().GetMethod('BeginUndoBatch',$flags).Invoke($search,@()) | Out-Null
    $state.GetType().GetMethod('SelectAll',$flags).Invoke($state,@($true)) | Out-Null
    if($checked.Count -ne 1) { throw 'Undo batch did not select only completed items.' }
    $script:batchUndo=0
    $search.add_UndoCheckedRequested({$script:batchUndo++})
    $applyButton=$search.Controls[0].GetControlFromPosition(0,1).Controls[2]
    if($null -eq $applyButton -or -not $applyButton.Enabled) { throw 'Undo batch action is unavailable.' }
    $applyButton.PerformClick()
    if($script:batchUndo -ne 1) { throw 'Undo batch action did not dispatch.' }
    $state.GetType().GetMethod('SetBatch',$flags).Invoke($state,@($false)) | Out-Null
    foreach($row in $state.GetType().GetField('Rows',$flags).GetValue($state)) {
        $rowKey=$assembly.GetType('ThesisAddIn.VSTO.Ui.IssueReviewState').GetMethod('Key',[Reflection.BindingFlags]'Static,NonPublic').Invoke($null,@($row))
        $state.GetType().GetField('Repaired',$flags).GetValue($state).Add($rowKey) | Out-Null
    }
    $search.ShowAnalysis((New-Object ThesisAddIn.Protocol.EngineAnalysisResult))
    $searchSummary=$search.GetType().GetField('_summary',$flags).GetValue($search).Text
    $suggestionSummary=$suggestion.GetType().GetField('_summary',$flags).GetValue($suggestion).Text
    if($searchSummary -ne $suggestionSummary) { throw 'Completed-item summaries diverged without a data rebuild.' }
    $suggestion.Height=210
    $suggestion.PerformLayout()
    $search.Height=210
    $search.PerformLayout()
    $leftViewport=$left.GetType().GetField('_viewport',$flags).GetValue($left)
    $leftViewport.GetType().GetMethod('ScrollTo',$flags).Invoke($leftViewport,@(17)) | Out-Null
    $third=$state.GetType().GetField('Rows',$flags).GetValue($state)[2]
    $state.GetType().GetMethod('Select',$flags).Invoke($state,@($third,$third.Location)) | Out-Null
    $rightBounds=$right.GetItemRectangle(2)
    if($rightBounds.Top -lt 0 -and $rightBounds.Height -le $right.Height) { throw 'Linked suggestion was not revealed.' }
    if($rightBounds.Top -ge $right.Height) { throw 'Linked suggestion remained outside the visible pane.' }
    if($leftViewport.GetType().GetProperty('Offset',$flags).GetValue($leftViewport) -ne 17) { throw 'Revealing a suggestion scrolled the search index.' }
    $bitmap=[Drawing.Bitmap]::new($right.Width,$right.Height)
    try { $right.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$right.Width,$right.Height)) } finally { $bitmap.Dispose() }
    Write-Output 'PASS: search cards dispatch solve and item undo; shared state and viewport preservation verified; compact cards render.'
} finally { $suggestion.Dispose(); $search.Dispose() }
