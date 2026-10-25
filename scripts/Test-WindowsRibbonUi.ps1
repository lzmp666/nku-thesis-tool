param([string]$AddinOutputDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) 'src\ThesisAddIn.VSTO\bin\Release'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$office = [Reflection.Assembly]::Load('Office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
$stdole = [Reflection.Assembly]::Load('stdole, Version=7.0.3300.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
$assembly = [Reflection.Assembly]::LoadFrom((Join-Path $AddinOutputDirectory 'ThesisAddIn.VSTO.dll'))
Add-Type -ReferencedAssemblies $office.Location,$stdole.Location -TypeDefinition @'
public sealed class UiProbe : Microsoft.Office.Core.IRibbonUI {
    public System.Collections.Generic.List<string> Changes = new System.Collections.Generic.List<string>();
    public void Invalidate() { throw new System.Exception("Whole Ribbon invalidation must not occur."); }
    public void InvalidateControl(string id) { Changes.Add(id); }
    public void InvalidateControlMso(string id) { Changes.Add(id); }
    public void ActivateTab(string id) {}
    public void ActivateTabMso(string id) {}
    public void ActivateTabQ(string id, string ns) {}
}
public sealed class ControlProbe : Microsoft.Office.Core.IRibbonControl {
    public string Id { get; set; }
    public string Tag { get { return ""; } }
    public object Context { get { return null; } }
    public static int Width(object picture) { return ((stdole.IPictureDisp)picture).Width; }
}
'@
$script:enabled = $false
$script:paneVisible = $false
$script:reads = 0
$state = [Func[bool]]{ $script:reads++; return $script:enabled }
$visibleCallback = [Func[bool]]{ $script:reads++; return $script:paneVisible }
$action = [Action]{}
$navigate = [Action[int]]{ param($direction) }
$show = [Action[bool]]{ param($value) }
$ribbon = [ThesisAddIn.VSTO.Ui.ThesisRibbon]::new($action,$action,$action,$action,$navigate,$navigate,$show,$visibleCallback,$state,$state,$state,$state,$state)
$ui = New-Object UiProbe
$ribbon.OnRibbonLoad($ui)
[xml]$deployment = Get-Content -LiteralPath (Join-Path $AddinOutputDirectory 'ThesisAddIn.VSTO.vsto') -Raw
$deploymentVersion = $deployment.SelectSingleNode('/*/*[local-name()="assemblyIdentity"]').GetAttribute('version')
$fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $AddinOutputDirectory 'ThesisAddIn.VSTO.dll')).FileVersion
if ($deploymentVersion -ne $fileVersion) { throw 'Deployment version and DLL file version differ; bump both before distributing an update.' }
$buildLabel = $ribbon.GetBuildLabel($null)
$information = $assembly.GetCustomAttributes([Reflection.AssemblyInformationalVersionAttribute], $false)[0].InformationalVersion
if ([string]::IsNullOrWhiteSpace($information) -or -not $buildLabel.Contains($information.Split('-')[0])) { throw 'The Ribbon does not expose its loaded assembly version.' }
Write-Output ('PASS: deployment/DLL version ' + $deploymentVersion + '; loaded Ribbon identity: ' + $buildLabel)
$ui.Changes.Clear()
$ribbon.Invalidate()
if ($ui.Changes.Count -ne 0) { throw 'Unchanged state invalidated controls.' }
$script:enabled = $true
$ribbon.Invalidate()
if ($ui.Changes.Count -ne 7) { throw 'Expected only the seven state-dependent buttons to invalidate.' }
$ui.Changes.Clear()
$before = $script:reads
[xml]$xml = $ribbon.GetCustomUI('Microsoft.Word.Document')
$buttons = $xml.SelectNodes('//*[@getImage="GetButtonImage"]')
if ($buttons.Count -ne 10) { throw 'Missing command icons.' }
foreach ($button in $buttons) {
    $control = New-Object ControlProbe
    $control.Id = [string]$button.id
    if ($null -eq $ribbon.GetButtonImage($control)) { throw ('Missing icon: '+$control.Id) }
    if ($button.getEnabled -and -not $ribbon.GetEnabled($control)) { throw 'Command state was not cached.' }
}
if ($script:reads -ne $before) { throw 'Ribbon getter queried delegates instead of cached state.' }
$script:paneVisible = $true
$ribbon.Invalidate()
if ($ui.Changes.Count -ne 1 -or $ui.Changes[0] -ne 'NkuShowIssues') { throw 'Visibility invalidated unrelated controls.' }
$brands = $xml.SelectNodes('//*[@getImage="GetBrandImage"]')
if ($brands.Count -ne 1 -or $brands[0].id -ne 'NkuEmblem' -or $brands[0].size -ne 'large') { throw 'Expected one readable large emblem and no campus group.' }
if ($xml.SelectNodes('//*[@id="NkuCampusStyle" or @id="NkuCampusPanorama"]').Count -ne 0) { throw 'Removed campus group remains.' }
foreach ($brand in $brands) {
    $control = New-Object ControlProbe
    $control.Id = [string]$brand.id
    $picture = $ribbon.GetBrandImage($control)
    if ($null -eq $picture -or [ControlProbe]::Width($picture) -le 0) { throw 'Brand image is invalid.' }
}
Write-Output 'PASS: no whole-Ribbon invalidation; unchanged state does not repaint; getters read cached state; open/close changes only the toggle; 10 command icons and one large emblem are valid.'
