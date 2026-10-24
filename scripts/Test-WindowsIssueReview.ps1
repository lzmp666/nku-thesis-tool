param([Parameter(Mandatory=$true)][string]$AddinOutputDirectory)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
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
    $state.GetType().GetMethod('SetBatch',$flags).Invoke($state,@($true)) | Out-Null
    $checked=$state.GetType().GetField('Checked',$flags).GetValue($state)
    $click=$left.GetType().GetMethod('OnMouseDown',$flags)
    $click.Invoke($left,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,14,18,0))) | Out-Null
    if($checked.Count -ne 1) { throw 'Search checkbox did not toggle the shared selection.' }
    $click.Invoke($right,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,14,18,0))) | Out-Null
    if($checked.Count -ne 0) { throw 'Suggestion checkbox did not update the same selection.' }
    $state.GetType().GetMethod('SetBatch',$flags).Invoke($state,@($false)) | Out-Null
    $script:repairs=0; $script:undos=0
    $suggestion.add_RepairRequested({$script:repairs++})
    $suggestion.add_UndoRequested({$script:undos++})
    $bounds=$right.GetItemRectangle(0)
    $click.Invoke($right,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,30,$bounds.Bottom-18,0))) | Out-Null
    if($script:repairs -ne 1) { throw 'Solve is not connected to the repair event.' }
    $selected=$suggestion.SelectedDiagnostic
    $key=$assembly.GetType('ThesisAddIn.VSTO.Ui.IssueReviewState').GetMethod('Key',[Reflection.BindingFlags]'Static,NonPublic').Invoke($null,@($selected))
    $state.GetType().GetField('Repaired',$flags).GetValue($state).Add($key) | Out-Null
    $click.Invoke($right,@([Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,100,$bounds.Bottom-18,0))) | Out-Null
    if($script:undos -ne 1) { throw 'Undo is not connected to the item undo event.' }
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
    Write-Output 'PASS: both review regions share checkboxes; solve and item undo dispatch real events; compact cards render.'
} finally { $suggestion.Dispose(); $search.Dispose() }
