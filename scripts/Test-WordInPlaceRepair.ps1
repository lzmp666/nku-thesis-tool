#requires -Version 5.1
#requires -PSEdition Desktop
[CmdletBinding()]
param([ValidateSet('Debug', 'Release')][string]$Configuration = 'Release')

$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
$outputDirectory = Join-Path $repository ('TestResults\word-inplace-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($outputDirectory) | Out-Null

function Assert-Condition([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

# WPS can take over Word.Application while exposing a compatible object model.
# Reject that registration before activation, fixture creation, or repair work.
function Get-MicrosoftWordAutomationExecutable {
    $classKey = 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID'
    Assert-Condition (Test-Path -LiteralPath $classKey) 'Microsoft Word automation is not registered.'
    $classId = (Get-Item -LiteralPath $classKey).GetValue('')
    $serverKey = 'Registry::HKEY_CLASSES_ROOT\CLSID\' + $classId + '\LocalServer32'
    Assert-Condition (Test-Path -LiteralPath $serverKey) 'The Word automation server is not registered.'
    $command = [string](Get-Item -LiteralPath $serverKey).GetValue('')
    $executable = $null
    if ($command -match '^\s*"([^"]+)"') { $executable = $Matches[1] }
    elseif ($command -match '^\s*(.+?\.exe)(?:\s|$)') { $executable = $Matches[1] }
    Assert-Condition (-not [string]::IsNullOrWhiteSpace($executable)) 'The Word automation server path is invalid.'
    Assert-Condition ([IO.Path]::GetFileName($executable) -ieq 'WINWORD.EXE') (
        'Word.Application points to a different application: ' + $executable + '. Microsoft Word acceptance has NOT run.')
    Assert-Condition (Test-Path -LiteralPath $executable -PathType Leaf) 'The Microsoft Word executable is missing.'
    Assert-Condition ((Get-Item -LiteralPath $executable).VersionInfo.CompanyName -like 'Microsoft*') 'The registered executable is not published by Microsoft.'
    return [IO.Path]::GetFullPath($executable)
}
$expectedWordExecutable = Get-MicrosoftWordAutomationExecutable

function Assert-Response($Response, [string]$Operation) {
    Assert-Condition ($null -ne $Response -and $Response.Ok) ($Operation + ': ' + $Response.Error.Message)
}
function Get-SharedFileHash([string]$Path) {
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($algorithm.ComputeHash($stream)).Replace('-', '') }
    finally { $algorithm.Dispose(); $stream.Dispose() }
}
function Get-ToolMarkCount($Document) {
    $count = 0
    for ($index = 1; $index -le $Document.Bookmarks.Count; $index++) {
        if ($Document.Bookmarks.Item($index).Name.StartsWith('_NKU_MARK_', [StringComparison]::Ordinal)) { $count++ }
    }
    return $count
}
function Get-ExpectedToolMarkCount($Analysis) {
    $count = 0
    foreach ($diagnostic in $Analysis.Diagnostics) {
        if ($diagnostic.Occurrences.Count -gt 0) { $count += $diagnostic.Occurrences.Count }
        else { $count++ }
    }
    return $count
}
function Assert-ToolMarks($Document, $Analysis, [string]$Operation) {
    $expected = Get-ExpectedToolMarkCount $Analysis
    $actual = Get-ToolMarkCount $Document
    Assert-Condition ($actual -eq $expected) ($Operation + ': expected ' + $expected + ' source marks, found ' + $actual + '.')
    Assert-Condition ($Document.Comments.Count -eq 0) ($Operation + ': the tool must not create duplicate reply/resolve comments.')
}
function Write-Snapshot($Document, [string]$Name) {
    $path = Join-Path $outputDirectory ($Name + '.docx')
    [ThesisAddIn.VSTO.Engine.FlatOpcPackageWriter]::Write($Document.WordOpenXML, $path)
    return $path
}
function Get-Run($Document, [string]$Snapshot, [string]$Location) {
    return [WordSmokeInterop]::GetRun($Document, $Snapshot, $Location)
}
function Apply-ToOpenDocument($Document, [string]$Snapshot, $Diagnostic, [string]$Name) {
    $locations = $null
    $ruleId = $null
    $location = $null
    if ($null -ne $Diagnostic) {
        $ruleId = $Diagnostic.RuleId
        $location = $Diagnostic.Location
        $locations = [System.Collections.Generic.List[string]]::new()
        foreach ($occurrence in $Diagnostic.Occurrences) { $locations.Add($occurrence.Location) }
    }
    $plan = $client.Plan($Snapshot, $ruleId, $location, $locations)
    Assert-Response $plan 'Plan'
    Assert-Condition ($plan.Plan.Patches.Count -gt 0) 'Expected at least one repair patch.'
    $appliedPath = Join-Path $outputDirectory ($Name + '-applied.docx')
    $applied = $client.Apply($Snapshot, $appliedPath, $ruleId, $location, $locations)
    Assert-Response $applied 'Apply'
    Assert-Condition ($applied.Apply.Failures.Count -eq 0) 'Engine reported patch failures.'
    $annotatedPath = Join-Path $outputDirectory ($Name + '-annotated.docx')
    $refreshed = $client.Annotate($appliedPath, $annotatedPath)
    Assert-Response $refreshed 'Annotate'
    $syncError = $null
    $synchronized = [WordSmokeInterop]::Apply(
        $Document, $Snapshot, $applied.Apply.Applied, $refreshed.Result, [ref]$syncError)
    Assert-Condition $synchronized ('In-place repair failed: ' + $syncError)
    Assert-ToolMarks $Document $refreshed.Result ($Name + ' in-place repair')
    return $applied.Apply.Applied.Count
}

$interop = Get-ChildItem 'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\*\Microsoft.Office.Interop.Word.dll' |
    Sort-Object FullName -Descending | Select-Object -First 1
Assert-Condition ($null -ne $interop) 'The Word primary interop assembly is required.'
$addInDirectory = Join-Path $repository ('src\ThesisAddIn.VSTO\bin\' + $Configuration)
$addInAssembly = Join-Path $repository ('src\ThesisAddIn.VSTO\obj\' + $Configuration + '\ThesisAddIn.VSTO.dll')
if (-not (Test-Path -LiteralPath $addInAssembly -PathType Leaf)) {
    $addInAssembly = Join-Path $addInDirectory 'ThesisAddIn.VSTO.dll'
}
Assert-Condition (Test-Path -LiteralPath $addInAssembly -PathType Leaf) 'The VSTO assembly is missing; compile the add-in before running Word acceptance.'
# Windows PowerShell does not read the add-in's binding redirects. Resolve its
# bundled dependencies in this test process without changing PowerShell/Office settings.
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Reflection;
public static class WordSmokeDependencies
{
    private static string directory;
    public static void Attach(string path)
    {
        directory = path;
        AppDomain.CurrentDomain.AssemblyResolve += Resolve;
    }
    private static Assembly Resolve(object sender, ResolveEventArgs args)
    {
        string path = Path.Combine(directory, new AssemblyName(args.Name).Name + ".dll");
        return File.Exists(path) ? Assembly.LoadFrom(path) : null;
    }
}
'@
[WordSmokeDependencies]::Attach($addInDirectory)
Add-Type -Path $interop.FullName
Add-Type -Path (Join-Path $addInDirectory 'ThesisAddIn.Protocol.dll')
Add-Type -Path $addInAssembly
$netstandardFacade = [Reflection.Assembly]::Load('netstandard, Version=2.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51').Location
# PowerShell's conversion binder cannot cast Word's coclass interface. Let the
# CLR perform COM QueryInterface inside this narrow test adapter instead.
Add-Type -ReferencedAssemblies @($interop.FullName,
    $netstandardFacade,
    (Join-Path $addInDirectory 'ThesisAddIn.Protocol.dll'),
    $addInAssembly) -TypeDefinition @'
using System;
using System.Collections.Generic;
using ThesisAddIn.Protocol;
using ThesisAddIn.VSTO.Engine;
using Word = Microsoft.Office.Interop.Word;
public static class WordSmokeInterop
{
    public static void Mark(object document) { WordDocumentSynchronizer.MarkWorkingCopy((Word.Document)document); }
    public static bool IsWorkingCopy(object document) { return WordDocumentSynchronizer.IsWorkingCopy((Word.Document)document); }
    public static object GetRun(object document, string snapshot, string location)
    {
        Word.Range range;
        string error;
        if (!WordDocumentSynchronizer.TryGetRunRange((Word.Document)document, snapshot, location, out range, out error))
            throw new InvalidOperationException(error);
        return range;
    }
    public static bool Apply(object document, string snapshot, List<EnginePatch> patches,
        EngineAnalysisResult analysis, out string error)
    {
        return WordDocumentSynchronizer.ApplyRepairsAndRefresh((Word.Document)document, snapshot, patches, analysis, out error);
    }
    public static bool Refresh(object document, string snapshot, EngineAnalysisResult analysis, out string error)
    {
        return WordDocumentSynchronizer.RefreshAnnotations((Word.Document)document, snapshot, analysis, out error);
    }
}
'@
Add-Type -AssemblyName System.IO.Compression.FileSystem
$hostAssembly = Join-Path $repository ('src\ThesisAddIn.VSTO\ProtocolHost\bin\' + $Configuration + '\net10.0\ThesisEngine.ProtocolHost.dll')
$client = [ThesisAddIn.Protocol.EngineProtocolClient]::new(
    [ThesisAddIn.Protocol.EngineProcessOptions]::ForDotnet($hostAssembly, [TimeSpan]::FromSeconds(45)))

# Extend only a synthetic fixture: three different positions need the same Chinese font.
$fixture = Join-Path $repository 'tests\fixtures\synthetic\demo\known_issues_sample.docx'
$fixtureHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $fixture).Hash
$source = Join-Path $outputDirectory 'synthetic-source.docx'
Copy-Item -LiteralPath $fixture -Destination $source
$archive = [IO.Compression.ZipFile]::Open($source, [IO.Compression.ZipArchiveMode]::Update)
try {
    $entry = $archive.GetEntry('word/document.xml')
    $reader = [IO.StreamReader]::new($entry.Open())
    try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $namespaces = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    $namespaces.AddNamespace('w', $namespace)
    $body = $xml.SelectSingleNode('/w:document/w:body', $namespaces)
    $paragraph = $body.SelectSingleNode('w:p', $namespaces)
    $section = $body.SelectSingleNode('w:sectPr', $namespaces)
    foreach ($font in @('黑体', '楷体')) {
        $copy = $paragraph.CloneNode($true)
        [void]$copy.SelectSingleNode('w:r/w:rPr/w:rFonts', $namespaces).SetAttribute('eastAsia', $namespace, $font)
        [void]$body.InsertBefore($copy, $section)
    }
    $entry.Delete()
    $entry = $archive.CreateEntry('word/document.xml')
    $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
    try { $xml.Save($writer) } finally { $writer.Dispose() }
} finally { $archive.Dispose() }
$sourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash
$workingCopy = Join-Path $outputDirectory 'synthetic-working-copy.docx'
Copy-Item -LiteralPath $source -Destination $workingCopy
$workingHash = Get-SharedFileHash $workingCopy

$word = $null
$document = $null
$fullThesisDocument = $null
$repairedDocument = $null
$ownsApplication = $false
$oldAlerts = $null
try {
    $word = New-Object -ComObject Word.Application
    $activatedExecutable = [IO.Path]::GetFullPath((Join-Path ([string]$word.Path) 'WINWORD.EXE'))
    Assert-Condition ([string]::Equals($activatedExecutable, $expectedWordExecutable, [StringComparison]::OrdinalIgnoreCase)) (
        'The activated automation host is not the registered Microsoft Word application. Acceptance has NOT run.')
    $ownsApplication = $word.Documents.Count -eq 0
    $oldAlerts = $word.DisplayAlerts
    $word.DisplayAlerts = 0
    if ($ownsApplication) { $word.Visible = $false }

    # Exercise the complete synthetic thesis in real Word, including paragraphs
    # before and after tables. This verifies flattened OOXML locations without
    # saving or replacing the source file.
    $fullThesis = Join-Path $repository 'tests\fixtures\synthetic\full-thesis\synthetic_complete_thesis.docx'
    $fullThesisHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $fullThesis).Hash
    $fullThesisDocument = $word.Documents.Open($fullThesis)
    $fullThesisPath = $fullThesisDocument.FullName
    $fullThesisWindowCount = $word.Windows.Count
    $fullThesisSnapshot = Write-Snapshot $fullThesisDocument 'full-thesis-before-analysis'
    $fullThesisAnalysis = $client.Analyze($fullThesisSnapshot)
    Assert-Response $fullThesisAnalysis 'Full synthetic thesis analysis'
    Assert-Condition ($fullThesisAnalysis.Result.Diagnostics.Count -eq 3) 'The full thesis should contain exactly three planned formatting defects.'
    Assert-Condition (@($fullThesisAnalysis.Result.UnsupportedScopes | Where-Object { $_ -like '*tables*' }).Count -gt 0) 'Table-specific formatting must be disclosed as unsupported.'
    Assert-Condition (@($fullThesisAnalysis.Result.UnsupportedScopes | Where-Object { $_ -like '*caption*' }).Count -gt 0) 'Caption formatting must be disclosed as unsupported.'
    Assert-Condition (@($fullThesisAnalysis.Result.UnsupportedScopes | Where-Object { $_ -like '*images*' }).Count -gt 0) 'Image quality must be disclosed as unsupported.'
    $syncError = $null
    Assert-Condition ([WordSmokeInterop]::Refresh($fullThesisDocument, $fullThesisSnapshot, $fullThesisAnalysis.Result, [ref]$syncError)) ('Full thesis range synchronization failed: ' + $syncError)
    Assert-ToolMarks $fullThesisDocument $fullThesisAnalysis.Result 'Full synthetic thesis analysis'
    foreach ($diagnostic in $fullThesisAnalysis.Result.Diagnostics) {
        foreach ($occurrence in $diagnostic.Occurrences) {
            $range = Get-Run $fullThesisDocument $fullThesisSnapshot $occurrence.Location
            Assert-Condition (-not [string]::IsNullOrWhiteSpace($range.Text)) ('Full thesis location resolved to empty text: ' + $occurrence.Location)
        }
    }
    Assert-Condition ($fullThesisDocument.FullName -eq $fullThesisPath -and $word.Windows.Count -eq $fullThesisWindowCount) 'Full thesis analysis changed or replaced the open document.'
    Assert-Condition ((Get-FileHash -Algorithm SHA256 -LiteralPath $fullThesis).Hash -eq $fullThesisHash) 'Full thesis analysis saved or overwrote the source file.'
    $fullThesisDocument.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
    $fullThesisDocument = $null

    $document = $word.Documents.Open($workingCopy)
    $pathBefore = $document.FullName
    $countBefore = $word.Documents.Count
    $windowsBefore = $word.Windows.Count
    Assert-Condition ($document.Comments.Count -eq 0 -and (Get-ToolMarkCount $document) -eq 0) 'The initial synthetic document must be unannotated.'
    $snapshot = Write-Snapshot $document 'before-first-analysis'
    $initial = $client.Analyze($snapshot)
    Assert-Response $initial 'First analysis'
    $syncError = $null
    Assert-Condition ([WordSmokeInterop]::Refresh($document, $snapshot, $initial.Result, [ref]$syncError)) ('First in-place analysis failed: ' + $syncError)
    Assert-Condition ($document.FullName -eq $pathBefore -and $word.Documents.Count -eq $countBefore -and $word.Windows.Count -eq $windowsBefore) 'Initial analysis changed the open document or opened a new window.'
    Assert-ToolMarks $document $initial.Result 'First analysis'
    Assert-Condition ((Get-SharedFileHash $workingCopy) -eq $workingHash) 'Initial analysis saved the document automatically.'
    $snapshot = Write-Snapshot $document 'before-repeat-analysis'
    $repeated = $client.Analyze($snapshot)
    Assert-Response $repeated 'Repeat analysis'
    Assert-Condition ([WordSmokeInterop]::Refresh($document, $snapshot, $repeated.Result, [ref]$syncError)) ('Repeat in-place analysis failed: ' + $syncError)
    Assert-Condition ($document.FullName -eq $pathBefore -and $word.Documents.Count -eq $countBefore -and $word.Windows.Count -eq $windowsBefore) 'Repeated analysis changed the open document or opened a new window.'
    Assert-ToolMarks $document $repeated.Result 'Repeated analysis'
    Assert-Condition ((Get-SharedFileHash $workingCopy) -eq $workingHash) 'Repeated analysis saved the document automatically.'
    # Rejected positioning must not undo the author's preceding edit. An empty
    # custom undo record is absent from Word's stack, so Undo would affect it.
    $probeRange = Get-Run $document $snapshot $repeated.Result.Diagnostics[0].Location
    $oldBold = $probeRange.Font.Bold
    $probeBold = if ($oldBold -eq 0) { -1 } else { 0 }
    $probeRange.Font.Bold = $probeBold
    $invalid = [ThesisAddIn.Protocol.EngineAnalysisResult]::new()
    $invalid.Diagnostics.Add([ThesisAddIn.Protocol.EngineDiagnostic]@{
        RuleId = 'demo/font-size'; Location = 'word/document.xml#paragraph[0]/run[999999]'
    })
    Assert-Condition (-not [WordSmokeInterop]::Refresh($document, $snapshot, $invalid, [ref]$syncError)) 'Invalid positioning was accepted.'
    Assert-Condition ($probeRange.Font.Bold -eq $probeBold) 'Failed validation undid the author edit.'
    Assert-ToolMarks $document $repeated.Result 'Rejected-position validation'
    Assert-Condition ([bool]$document.Undo()) 'The author edit could not be undone after validation failed.'
    Assert-Condition ($probeRange.Font.Bold -eq $oldBold) 'Validation failure changed the undo stack.'
    $snapshot = Write-Snapshot $document 'before-single'
    $analysis = $client.Analyze($snapshot)
    Assert-Response $analysis 'Analyze'
    $group = @($analysis.Result.Diagnostics | Where-Object RuleId -eq 'demo/east-asian-font')[0]
    Assert-Condition ($group.Occurrences.Count -ge 3) 'Repeated font errors were not grouped.'
    $originalFonts = @{}
    foreach ($occurrence in $group.Occurrences) {
        $originalFonts[$occurrence.Location] = (Get-Run $document $snapshot $occurrence.Location).Font.NameFarEast
    }
    $marksBeforeSingle = Get-ToolMarkCount $document
    $singlePatchCount = Apply-ToOpenDocument $document $snapshot $group 'single'
    $marksAfterSingle = Get-ToolMarkCount $document
    Assert-Condition ($singlePatchCount -eq $group.Occurrences.Count) 'Single-item repair did not cover every grouped occurrence.'
    foreach ($occurrence in $group.Occurrences) {
        Assert-Condition ((Get-Run $document $snapshot $occurrence.Location).Font.NameFarEast -eq '宋体') 'A grouped font target was not repaired.'
    }
    Assert-Condition ([bool]$document.Undo()) 'Word could not undo the repair.'
    foreach ($occurrence in $group.Occurrences) {
        Assert-Condition ((Get-Run $document $snapshot $occurrence.Location).Font.NameFarEast -eq $originalFonts[$occurrence.Location]) 'One undo did not restore all grouped fonts.'
    }
    Assert-Condition ((Get-ToolMarkCount $document) -eq $marksBeforeSingle) 'One undo did not restore source marks.'
    Assert-Condition ([bool]$document.Redo()) 'Word could not redo the repair.'
    Assert-Condition ((Get-ToolMarkCount $document) -eq $marksAfterSingle) 'One redo did not restore source marks.'
    foreach ($occurrence in $group.Occurrences) {
        Assert-Condition ((Get-Run $document $snapshot $occurrence.Location).Font.NameFarEast -eq '宋体') 'One redo did not restore the grouped repair.'
    }

    $bulkSnapshot = Write-Snapshot $document 'before-bulk'
    $bulkPlan = $client.Plan($bulkSnapshot, $null, $null)
    Assert-Response $bulkPlan 'Bulk plan'
    Assert-Condition ($bulkPlan.Plan.Patches.Count -eq 1 -and $bulkPlan.Plan.Patches[0].RuleId -eq 'demo/font-size') 'The synthetic bulk plan has unexpected patches.'
    $sizeLocation = $bulkPlan.Plan.Patches[0].Location
    $sizeBefore = (Get-Run $document $bulkSnapshot $sizeLocation).Font.Size
    $marksBeforeBulk = Get-ToolMarkCount $document
    $bulkPatchCount = Apply-ToOpenDocument $document $bulkSnapshot $null 'bulk'
    $marksAfterBulk = Get-ToolMarkCount $document
    Assert-Condition ((Get-Run $document $bulkSnapshot $sizeLocation).Font.Size -eq 12) 'Bulk repair did not change the font size.'
    Assert-Condition ([bool]$document.Undo()) 'Word could not undo the bulk repair.'
    Assert-Condition ((Get-Run $document $bulkSnapshot $sizeLocation).Font.Size -eq $sizeBefore) 'One undo did not restore the bulk formatting.'
    Assert-Condition ((Get-ToolMarkCount $document) -eq $marksBeforeBulk) 'One undo did not restore bulk source marks.'
    Assert-Condition ([bool]$document.Redo()) 'Word could not redo the bulk repair.'
    Assert-Condition ((Get-ToolMarkCount $document) -eq $marksAfterBulk) 'One redo did not restore bulk source marks.'
    Assert-Condition ((Get-Run $document $bulkSnapshot $sizeLocation).Font.Size -eq 12) 'One redo did not restore the bulk repair.'
    Assert-Condition ($document.FullName -eq $pathBefore -and $word.Documents.Count -eq $countBefore) 'Repair opened or replaced a document.'
    Assert-Condition (-not $document.Saved) 'Word did not retain unsaved edits.'
    Assert-Condition ((Get-SharedFileHash $workingCopy) -eq $workingHash) 'Repair wrote to disk before author save.'
    $document.Save()
    Assert-Condition ($document.Saved) 'Explicit author save failed.'
    $after = $client.Analyze((Write-Snapshot $document 'after-save'))
    Assert-Response $after 'Analyze after save'
    Assert-Condition (@($after.Result.Diagnostics | Where-Object FixLevel -eq 'auto').Count -eq 0) 'Auto-fixable issues remain.'
    $remaining = $after.Result.Diagnostics.Count
    # Keep one test document open while reopening the clean copy. A non-user-
    # controlled Word instance may shut down when its last document is closed.
    $repairedDocument = $document

    # A clean work copy has no source marks. Its identity must survive save/reopen.
    $cleanInput = Join-Path $repository 'tests\fixtures\synthetic\demo\clean_sample.docx'
    $cleanCopy = Join-Path $outputDirectory 'synthetic-clean-working-copy.docx'
    $clean = $client.Annotate($cleanInput, $cleanCopy)
    Assert-Response $clean 'Annotate clean fixture'
    $document = $word.Documents.Open($cleanCopy)
    [WordSmokeInterop]::Mark($document)
    $document.Save()
    $document.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
    $document = $word.Documents.Open($cleanCopy)
    Assert-Condition ([WordSmokeInterop]::IsWorkingCopy($document)) 'Clean working-copy identity did not persist.'
    Assert-Condition ($document.Comments.Count -eq 0 -and (Get-ToolMarkCount $document) -eq 0) 'Clean document has tool annotations.'
    $syncError = $null
    Assert-Condition ([WordSmokeInterop]::Refresh(
        $document, (Write-Snapshot $document 'clean-reopened'), $clean.Result, [ref]$syncError)) ('Clean reanalysis failed: ' + $syncError)
    Assert-Condition ((Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash -eq $sourceHash) 'Synthetic source was overwritten.'
    Assert-Condition ((Get-FileHash -Algorithm SHA256 -LiteralPath $fixture).Hash -eq $fixtureHash) 'Repository fixture was overwritten.'
    [ordered]@{
        status = 'passed'
        fullThesisWordMapping = 'passed'
        fullThesisDiagnostics = $fullThesisAnalysis.Result.Diagnostics.Count
        initialAndRepeatedAnalysisInSameWindow = 'passed'
        groupedOccurrences = $group.Occurrences.Count
        singleRepairPatches = $singlePatchCount
        bulkRepairPatches = $bulkPatchCount
        sourceMarkBookmarks = Get-ToolMarkCount $repairedDocument
        remainingManualOrSuggestIssues = $remaining
        undoRedo = 'passed'
        sameOpenDocument = 'passed'
        explicitSaveOnly = 'passed'
        cleanCopyIdentityAfterReopen = 'passed'
        sourceHashesUnchanged = $true
    } | ConvertTo-Json
} catch {
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if ($null -ne $fullThesisDocument) {
        try { $fullThesisDocument.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges) } catch { }
    }
    if ($null -ne $document) {
        try { $document.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges) } catch { }
    }
    if ($null -ne $repairedDocument) {
        try { $repairedDocument.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges) } catch { }
    }
    if ($null -ne $word) {
        if ($null -ne $oldAlerts) { try { $word.DisplayAlerts = $oldAlerts } catch { } }
        if ($ownsApplication) {
            try { $word.Quit([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges) } catch { }
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
