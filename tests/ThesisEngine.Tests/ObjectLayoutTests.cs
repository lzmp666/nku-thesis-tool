using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class ObjectLayoutTests
{
    [Fact]
    public void CompleteSyntheticThesisRepairsBodyCaptionsAndOversizedObjectsAndPreservesSource()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln"))) root = root.Parent;
        Assert.NotNull(root);
        var input = Path.Combine(root.FullName, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_object_layout_errors.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var source = File.ReadAllBytes(input);
        try
        {
            var analysis = Analyze(input);
            Assert.Equal(11, analysis.Diagnostics.Count);
            Assert.Equal(2, WidthPlan(input).Length);
            var result = PatchEngine.ApplyPatches(input, PatchEngine.PlanPatches(analysis), output);
            Assert.True(result.Succeeded, string.Join(";", result.Failures.Select(f => f.Reason)));
            Assert.Empty(Analyze(output).Diagnostics);
            using (var document = WordprocessingDocument.Open(output, false))
            {
                var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(document).Select(error => error.Description + " at " + error.Path?.XPath).ToArray();
                Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
            }
            var before = ReadXml(input); var after = ReadXml(output);
            Assert.Equal(before.Descendants(W + "t").Select(e => e.Value), after.Descendants(W + "t").Select(e => e.Value));
            Assert.Equal(before.Descendants(W + "tbl").Count(), after.Descendants(W + "tbl").Count());
            using var first = ZipFile.OpenRead(input); using var second = ZipFile.OpenRead(output);
            foreach (var part in first.Entries.Where(e => e.FullName != "word/document.xml"))
                Assert.Equal(Bytes(part), Bytes(second.GetEntry(part.FullName)!));
            Assert.Equal(source, File.ReadAllBytes(input));
        }
        finally { File.Delete(output); }
    }

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string Section = "<w:sectPr><w:pgSz w:w=\"12000\" w:h=\"17000\"/><w:pgMar w:left=\"1800\" w:right=\"1800\" w:gutter=\"0\"/></w:sectPr>";
    private const string Picture = """
        <w:p><w:r><w:drawing><wp:inline><wp:extent cx="7620000" cy="3810000"/>
        <wp:docPr id="1" name="synthetic picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="7620000" cy="3810000"/></a:xfrm>
        <a:extLst><a:ext uri="keep-me"/></a:extLst></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
        """;
    private static string Caption(string text) => $"<w:p><w:pPr><w:pStyle w:val=\"Caption\"/><w:jc w:val=\"center\"/><w:spacing w:before=\"120\" w:after=\"240\" w:line=\"240\"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia=\"宋体\"/><w:sz w:val=\"21\"/></w:rPr><w:t>{text}</w:t></w:r></w:p>";
    private static string Table => "<w:tbl><w:tblPr><w:tblW w:type=\"dxa\" w:w=\"12000\"/><w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/><w:gridCol w:w=\"8000\"/></w:tblGrid>" +
        string.Concat(Enumerable.Range(0, 3).Select(i => $"<w:tr><w:tc><w:tcPr><w:tcW w:type=\"dxa\" w:w=\"4000\"/></w:tcPr><w:p><w:r><w:t>项目 {i}</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type=\"dxa\" w:w=\"8000\"/></w:tcPr><w:p><w:r><w:t>value {i}</w:t></w:r></w:p></w:tc></w:tr>")) + "</w:tbl>";

    [Fact]
    public void RepairsPictureAndTableWidthsWithoutChangingAnyTextMediaOrOtherParts()
    {
        var input = Package(Picture + Caption("图1.1　合成图") + Caption("表1.1　合成表") + Table);
        var output = input + ".fixed.docx";
        try
        {
            var before = File.ReadAllBytes(input);
            var plan = WidthPlan(input);
            Assert.Equal(2, plan.Length);
            Assert.All(plan, p => Assert.Equal("420pt", p.DesiredAfter));
            var originalXml = ReadXml(input);
            var result = PatchEngine.ApplyPatches(input, plan, output);
            Assert.True(result.Succeeded, string.Join(";", result.Failures.Select(f => f.Reason)));
            var repaired = ReadXml(output);
            Assert.Equal(originalXml.Descendants(W + "t").Select(e => e.Value), repaired.Descendants(W + "t").Select(e => e.Value));
            Assert.Equal("5334000", repaired.Descendants(WP + "extent").Single().Attribute("cx")!.Value);
            Assert.Equal("2667000", repaired.Descendants(WP + "extent").Single().Attribute("cy")!.Value);
            Assert.Equal("5334000", repaired.Descendants(A + "xfrm").Single().Element(A + "ext")!.Attribute("cx")!.Value);
            Assert.Single(repaired.Descendants(A + "ext").Single(e => (string?)e.Attribute("uri") == "keep-me").Attributes());
            Assert.Equal(new[] { "2800", "5600" }, repaired.Descendants(W + "gridCol").Select(e => (string)e.Attribute(W + "w")!));
            Assert.Equal("8400", repaired.Descendants(W + "tblW").Single().Attribute(W + "w")!.Value);
            Assert.Empty(WidthPlan(output));
            Assert.Equal(before, File.ReadAllBytes(input));
            using var source = ZipFile.OpenRead(input); using var destination = ZipFile.OpenRead(output);
            foreach (var entry in source.Entries.Where(e => e.FullName != "word/document.xml"))
                Assert.Equal(Bytes(entry), Bytes(destination.GetEntry(entry.FullName)!));
        }
        finally { File.Delete(input); File.Delete(output); }
    }

    [Fact]
    public void RejectsChangedPageWidthAndPublishesNoPartialOutput()
    {
        var input = Package(Picture + Caption("图1.1　合成图")); var output = input + ".fixed.docx";
        try
        {
            var patch = Assert.Single(WidthPlan(input)) with { DesiredAfter = "400pt" };
            Assert.False(PatchEngine.ApplyPatches(input, new[] { patch }, output).Succeeded);
            Assert.False(File.Exists(output));
        }
        finally { File.Delete(input); File.Delete(output); }
    }

    [Theory]
    [InlineData("floating")]
    [InlineData("multiple")]
    [InlineData("columns")]
    [InlineData("no-page-setup")]
    [InlineData("rotated")]
    [InlineData("unpaired")]
    [InlineData("smaller")]
    public void UncertainOrAlreadyFittingPicturesAreNeverResized(string kind)
    {
        var picture = kind switch
        {
            "floating" => Picture.Replace("wp:inline", "wp:anchor"),
            "multiple" => Picture.Replace("</w:r>", Picture[5..^6] + "</w:r>"),
            "rotated" => Picture.Replace("<a:xfrm>", "<a:xfrm rot=\"5400000\">"),
            "smaller" => Picture.Replace("7620000", "2540000").Replace("3810000", "1270000"),
            _ => Picture
        };
        var section = kind == "columns" ? Section.Replace("</w:sectPr>", "<w:cols w:num=\"2\"/></w:sectPr>") : kind == "no-page-setup" ? "<w:sectPr/>" : Section;
        var path = Package(picture + (kind == "unpaired" ? "<w:p><w:r><w:t>中间的正文</w:t></w:r></w:p>" : "") + Caption("图1.1　合成图"), section);
        try { Assert.Empty(WidthPlan(path)); } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("merged")]
    [InlineData("autofit")]
    [InlineData("style")]
    [InlineData("different-cell-width")]
    public void ComplexOrInheritedTablesAreNotGivenAutomaticGeometryPatches(string kind)
    {
        var table = kind switch
        {
            "merged" => Table.Replace("<w:tcPr>", "<w:tcPr><w:gridSpan w:val=\"2\"/>"),
            "autofit" => Table.Replace("w:type=\"fixed\"", "w:type=\"autofit\""),
            "style" => Table.Replace("<w:tblPr>", "<w:tblPr><w:tblStyle w:val=\"Grid\"/>"),
            _ => Table.Replace("<w:tcW w:type=\"dxa\" w:w=\"4000\"", "<w:tcW w:type=\"dxa\" w:w=\"3000\"")
        };
        var path = Package(Caption("表1.1　合成表") + table);
        try { Assert.Empty(WidthPlan(path)); } finally { File.Delete(path); }
    }

    [Fact]
    public void UsesTheObjectsOwnSectionInsteadOfTheLastSectionOfTheDocument()
    {
        var sectionBreak = "<w:p><w:pPr>" + Section.Replace("<w:sectPr>", "<w:sectPr><w:type w:val=\"nextPage\"/>") + "</w:pPr></w:p>";
        var path = Package(Picture + Caption("图1.1　合成图") + sectionBreak, Section.Replace("12000", "20000"));
        try { Assert.Equal("420pt", Assert.Single(WidthPlan(path)).DesiredAfter); } finally { File.Delete(path); }
    }

    [Fact]
    public void EnglishCaptionHasNoInventedChineseGapRequirement()
    {
        var path = Package(Picture + Caption("Figure 1.1 Synthetic example"));
        try { Assert.DoesNotContain(Analyze(path).Diagnostics, d => d.RuleId == FontSizeAnalyzer.RuleIdCaptionNumberGap); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void PictureLimitIncludesInheritedParagraphIndentsAndRejectsCharacterBasedIndentation()
    {
        var styles = $"<w:styles xmlns:w=\"{W}\"><w:style w:type=\"paragraph\" w:styleId=\"Base\"><w:pPr><w:ind w:left=\"720\"/></w:pPr></w:style><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:basedOn w:val=\"Base\"/><w:pPr><w:ind w:right=\"360\"/></w:pPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"Caption\"/></w:styles>";
        var path = Package(Picture + Caption("图1.1　合成图"), styles: styles);
        var unsupported = Package(Picture + Caption("图1.1　合成图"), styles: styles.Replace("w:left=\"720\"", "w:leftChars=\"200\""));
        try
        {
            Assert.Equal("366pt", Assert.Single(WidthPlan(path)).DesiredAfter);
            Assert.Empty(WidthPlan(unsupported));
        }
        finally { File.Delete(path); File.Delete(unsupported); }
    }

    private static AnalysisResult Analyze(string path) => new FontSizeAnalyzer().Analyze(new WordFormattingReader().Read(path), DemoRuleSet.NkuDraftSample);
    private static Patch[] WidthPlan(string path) => PatchEngine.PlanPatches(Analyze(path)).Patches.Where(p => p.RuleId is FontSizeAnalyzer.RuleIdPictureWidth or FontSizeAnalyzer.RuleIdTableWidth).ToArray();
    private static string Package(string body, string section = Section, string? styles = null)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Add(string name, string text) { using var writer = new StreamWriter(zip.CreateEntry(name).Open()); writer.Write(text); }
        Add("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
        Add("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        Add("word/document.xml", $"<w:document xmlns:w=\"{W}\" xmlns:wp=\"{WP}\" xmlns:a=\"{A}\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>{body}{section}</w:body></w:document>");
        Add("word/styles.xml", styles ?? $"<w:styles xmlns:w=\"{W}\"><w:style w:type=\"paragraph\" w:styleId=\"Caption\"/></w:styles>");
        Add("word/media/synthetic.bin", "opaque-image-bytes");
        Add("word/_rels/document.xml.rels", "keep-original-relationships");
        return path;
    }
    private static XDocument ReadXml(string path) { using var z = ZipFile.OpenRead(path); using var s = z.GetEntry("word/document.xml")!.Open(); return XDocument.Load(s); }
    private static byte[] Bytes(ZipArchiveEntry entry) { using var stream = entry.Open(); using var memory = new MemoryStream(); stream.CopyTo(memory); return memory.ToArray(); }
}
