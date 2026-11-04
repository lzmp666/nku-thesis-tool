param([string]$AddinOutputDirectory)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$output = [IO.Path]::GetFullPath($AddinOutputDirectory)
# PowerShell does not read the VSTO DLL's app.config binding redirects. Mirror
# the installed add-in's JSON redirect in this isolated component-test process.
Add-Type -TypeDefinition @'
public static class OutlineProbeAssemblyResolver {
    public static string Directory;
    public static void Register() { System.AppDomain.CurrentDomain.AssemblyResolve += Resolve; }
    public static System.Reflection.Assembly Resolve(object sender, System.ResolveEventArgs args) {
        var name = new System.Reflection.AssemblyName(args.Name).Name;
        if (name != "System.Text.Json" && name != "System.Text.Encodings.Web") return null;
        return System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(Directory, name + ".dll"));
    }
}
'@
[OutlineProbeAssemblyResolver]::Directory = $output
[OutlineProbeAssemblyResolver]::Register()
foreach ($dependency in @('System.Runtime.CompilerServices.Unsafe.dll','System.Memory.dll','System.Buffers.dll',
    'System.Threading.Tasks.Extensions.dll','Microsoft.Bcl.AsyncInterfaces.dll','System.Text.Encodings.Web.dll','System.Text.Json.dll')) {
    [void][Reflection.Assembly]::LoadFrom((Join-Path $output $dependency))
}
[void][Reflection.Assembly]::LoadFrom((Join-Path $output 'ThesisAddIn.Protocol.dll'))
$assembly = [Reflection.Assembly]::LoadFrom((Join-Path $output 'ThesisAddIn.VSTO.dll'))
$flags = [Reflection.BindingFlags]'Instance,NonPublic'
$type = $assembly.GetType('ThesisAddIn.VSTO.Ui.ThesisOutlineReviewDialog', $true)
$review = New-Object ThesisAddIn.Protocol.EngineStructureReviewResult
function Add-Block([string]$Id, [string]$Section, [double]$Confidence, [bool]$NeedsReview, [bool]$Table) {
    $block = New-Object ThesisAddIn.Protocol.EngineStructureBlock
    $block.BlockId = $Id; $block.Section = $Section; $block.Kind = 'Paragraph'
    $block.ExactText = '合成测试段落 ' + $Id; $block.Location = 'word/document.xml#paragraph[' + $review.Blocks.Count + ']'
    $block.Confidence = $Confidence; $block.RequiresReview = $NeedsReview
    $block.HasCompleteTextMapping = $true; $block.CanConfirm = -not $Table; $block.IsTableCell = $Table
    $review.Blocks.Add($block)
}
Add-Block 'body' 'Body' 0.95 $false $false
Add-Block 'abstract' 'ChineseAbstract' 0.85 $true $false
Add-Block 'table' 'Body' 1.0 $false $true
Add-Block 'unknown' 'Unknown' 0.5 $true $false
$constructor = $type.GetConstructors($flags) | Where-Object { $_.GetParameters().Count -eq 1 } | Select-Object -First 1
if ($null -eq $constructor) { throw 'Outline dialog review constructor is missing.' }
$dialog = $constructor.Invoke([object[]]@($review.PSObject.BaseObject))
try {
    $grid = $type.GetField('_grid', $flags).GetValue($dialog)
    if ($grid.Rows.Count -ne 4) { throw 'Outline rows were lost.' }
    if (@($grid.Rows | Where-Object { $_.Cells['confirmed'].Value }).Count -ne 0) { throw 'Candidates were implicitly approved.' }
    [void]$type.GetMethod('ConfirmRuleCandidates', $flags).Invoke($dialog, @())
    if (-not $grid.Rows[0].Cells['confirmed'].Value -or $grid.Rows[1].Cells['confirmed'].Value -or
        $grid.Rows[2].Cells['confirmed'].Value -or $grid.Rows[3].Cells['confirmed'].Value) { throw 'Bulk rule approval included uncertain/physical/unknown blocks.' }
    $grid.Rows[0].Cells['kind'].Value = 'Heading'
    if ($grid.Rows[0].Cells['confirmed'].Value) { throw 'Role edits retained old approval.' }
    $grid.Rows[0].Cells['kind'].Value = 'Paragraph'
    $grid.Rows[0].Cells['confirmed'].Value = $true
    $grid.Rows[1].Cells['confirmed'].Value = $true
    $grid.Rows[2].Cells['confirmed'].Value = $true
    $grid.Rows[3].Cells['confirmed'].Value = $true
    if ($grid.Rows[2].Cells['confirmed'].Value -or $grid.Rows[3].Cells['confirmed'].Value) { throw 'Unsupported blocks can be approved.' }
    [void]$type.GetMethod('AcceptApprovals', $flags).Invoke($dialog, @())
    if ($type.GetProperty('ConfirmedRoles', $flags).GetValue($dialog).Count -ne 0) { throw 'Missing draft acknowledgment was accepted.' }
    $type.GetField('_acknowledgment', $flags).GetValue($dialog).Checked = $true
    [void]$type.GetMethod('AcceptApprovals', $flags).Invoke($dialog, @())
    $approved = $type.GetProperty('ConfirmedRoles', $flags).GetValue($dialog)
    if ($approved.Count -ne 2 -or $approved[0].BlockId -ne 'body' -or $approved[1].BlockId -ne 'abstract') {
        throw 'Explicit author selections did not round-trip.'
    }
    Write-Output 'PASS: outline starts unapproved, bulk rules omit uncertain/unsafe roles, edits revoke approval, draft acknowledgment is mandatory, exact selected IDs returned.'
} finally { $dialog.Dispose() }
