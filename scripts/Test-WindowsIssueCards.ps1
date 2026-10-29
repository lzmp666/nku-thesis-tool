param([string]$AddinOutputDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) 'src\ThesisAddIn.VSTO\bin\Release'))
$ErrorActionPreference = 'Stop'
# Render isolated controls into bitmaps; no Word COM or user app automation.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Reflection.Assembly]::LoadFrom((Join-Path $AddinOutputDirectory 'ThesisAddIn.Protocol.dll')) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path $AddinOutputDirectory 'ThesisAddIn.VSTO.dll')) | Out-Null
$pane = New-Object ThesisAddIn.VSTO.Ui.DiagnosticsPane
try {
    $pane.Size = New-Object Drawing.Size(390,720)
    $pane.CreateControl()
    $analysis = New-Object ThesisAddIn.Protocol.EngineAnalysisResult
    for ($index=0; $index -lt 9; $index++) {
        $diagnostic = New-Object ThesisAddIn.Protocol.EngineDiagnostic
        $diagnostic.RuleId = 'demo/font-size'
        $diagnostic.Expected = ([string](12+$index)+'pt')
        $diagnostic.Actual = '10.5pt'
        $diagnostic.Location = ('word/document.xml#paragraph['+$index+']/run[0]')
        $diagnostic.FixLevel = 'auto'
        $analysis.Diagnostics.Add($diagnostic)
    }
    $pane.ShowAnalysis($analysis)
    $flags = [Reflection.BindingFlags]'Instance,NonPublic'
    $list = $pane.GetType().GetField('_issueList',$flags).GetValue($pane)
    $viewport = $list.GetType().GetField('_viewport',$flags).GetValue($list)
    $getOffset = $viewport.GetType().GetProperty('Offset',$flags)
    $scroll = $viewport.GetType().GetMethod('ScrollTo',$flags)
    $scroll.Invoke($viewport,@(47)) | Out-Null
    $events = [Collections.Generic.List[int]]::new()
    $list.add_SelectedIndexChanged({ $events.Add($list.SelectedIndex) })
    for ($frame=0; $frame -lt 12; $frame++) {
        $list.SelectedIndex = $frame % 2
        if ($getOffset.GetValue($viewport) -ne 47) { throw 'Selection snapped the partial-card scroll offset.' }
        $bitmap = New-Object Drawing.Bitmap($list.Width,$list.Height)
        try {
            $list.DrawToBitmap($bitmap, [Drawing.Rectangle]::new(0,0,$list.Width,$list.Height))
            $dark = 0; $blue = 0
            for ($y=5; $y -lt [Math]::Min(220,$bitmap.Height); $y++) {
                for ($x=5; $x -lt $list.ContentWidth-5; $x++) {
                    $pixel = $bitmap.GetPixel($x,$y)
                    if ($pixel.R -lt 170 -and $pixel.G -lt 170 -and $pixel.B -lt 170) { $dark++ }
                    if ($pixel.R -lt 35 -and $pixel.G -gt 85 -and $pixel.G -lt 165 -and $pixel.B -gt 175) { $blue++ }
                }
            }
            if ($dark -lt 100) { throw 'Card text disappeared during a selected-frame render.' }
            if ($blue -gt 0) { throw 'Native blue ListBox selection pixels leaked into a frame.' }
        } finally { $bitmap.Dispose() }
    }
    $before = $events.Count
    $list.SelectedIndex = $list.SelectedIndex
    if ($events.Count -ne $before) { throw 'Selecting an unchanged card repeated selection events.' }
    if ($list.AccessibilityObject.GetChildCount() -ne 9) { throw 'Accessible issue cards are missing.' }
    Write-Output 'PASS: 12 complete selected frames retain text and contain no native blue selection fill; selection preserves a 47px partial-card offset; repeated selection is silent; 9 accessible cards remain.'
} finally { $pane.Dispose() }
