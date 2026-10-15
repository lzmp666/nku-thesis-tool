using System.IO.Compression;
using System.Xml.Linq;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class PatchEngineTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void PlansOnlyAutoFindingsAndKeepsDeferredDiagnostics()
    {
        var result = new AnalysisResult(
            [
                new Diagnostic(FontSizeAnalyzer.RuleIdEastAsiaFont, "warning", "word/document.xml#paragraph[0]/run[0]", "宋体", "黑体", "auto", "font"),
                new Diagnostic(FontSizeAnalyzer.RuleIdAsciiFont, "warning", "word/document.xml#paragraph[0]/run[0]", "Times New Roman", "Arial", "suggest", "font"),
                new Diagnostic(FontSizeAnalyzer.RuleIdFontSize, "warning", "word/document.xml#paragraph[0]/run[0]", "12pt", "(unknown)", "manual", "size")
            ],
            "demo-0.1.0", "covered", []);

        var plan = PatchEngine.PlanPatches(result);

        var patch = Assert.Single(plan.Patches);
        Assert.Equal("黑体", patch.ExpectedBefore);
        Assert.Equal("宋体", patch.DesiredAfter);
        Assert.Equal(2, plan.DeferredDiagnostics.Count);
    }

    [Fact]
    public void AppliesBatchWithoutChangingTextOrInputAndIsIdempotent()
    {
        var input = CreatePackage("黑体", "Arial", "22", "正文内容");
        var output = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.fixed.docx");
        try
        {
            var beforeBytes = File.ReadAllBytes(input);
            var before = new WordFormattingReader().Read(input);
            var analysis = new FontSizeAnalyzer().Analyze(before, DemoRuleSet.NkuDraftSample);
            var plan = PatchEngine.PlanPatches(analysis);
            var result = PatchEngine.ApplyPatches(input, plan, output);

            Assert.True(result.Succeeded);
            Assert.Equal(beforeBytes, File.ReadAllBytes(input));
            var after = new WordFormattingReader().Read(output);
            var run = Assert.Single(after.Paragraphs).Runs.Single();
            Assert.Equal("宋体", run.EastAsiaFont);
            Assert.Equal(12, run.FontSizePt);
            Assert.Equal("正文内容", run.Text);
            var second = new FontSizeAnalyzer().Analyze(after, DemoRuleSet.NkuDraftSample);
            Assert.Empty(PatchEngine.PlanPatches(second).Patches);
            Assert.Single(second.Diagnostics); // western font remains suggest
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    [Fact]
    public void InsertsInheritedFontAndSizeInSchemaOrder()
    {
        var input = CreateInheritedPackage();
        var output = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.fixed.docx");
        try
        {
            var originalBytes = File.ReadAllBytes(input);
            var before = new FontSizeAnalyzer().Analyze(new WordFormattingReader().Read(input), DemoRuleSet.NkuDraftSample);
            var plan = PatchEngine.PlanPatches(before);
            Assert.Equal(2, plan.Patches.Count);

            var result = PatchEngine.ApplyPatches(input, plan, output);

            Assert.True(result.Succeeded);
            Assert.Equal(originalBytes, File.ReadAllBytes(input));
            var after = new WordFormattingReader().Read(output);
            var run = Assert.Single(Assert.Single(after.Paragraphs).Runs);
            Assert.Equal("宋体", run.EastAsiaFont);
            Assert.Equal(12, run.FontSizePt);

            using var archive = ZipFile.OpenRead(output);
            using var documentStream = archive.GetEntry("word/document.xml")!.Open();
            var document = XDocument.Load(documentStream);
            var properties = Assert.Single(document.Descendants(W + "rPr"));
            Assert.Equal(new[] { "rStyle", "rFonts", "sz", "highlight" },
                properties.Elements().Select(element => element.Name.LocalName));
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    [Fact]
    public void StalePreconditionReturnsFailureAndPublishesNoOutput()
    {
        var input = CreatePackage("黑体", "Arial", "22", "正文内容");
        var output = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.fixed.docx");
        try
        {
            var stale = new Patch(FontSizeAnalyzer.RuleIdEastAsiaFont,
                "word/document.xml#paragraph[0]/run[0]", "不存在的字体", "宋体", "test");
            var result = PatchEngine.ApplyPatches(input, new PatchPlan([stale]), output);

            Assert.False(result.Succeeded);
            Assert.Single(result.Failures);
            Assert.False(File.Exists(output));
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    [Fact]
    public void BatchPreconditionFailurePublishesNoOutputAndDoesNotChangeInput()
    {
        var input = CreatePackage("黑体", "Arial", "22", "正文内容");
        var output = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.fixed.docx");
        try
        {
            var originalBytes = File.ReadAllBytes(input);
            var patches = new[]
            {
                new Patch(FontSizeAnalyzer.RuleIdEastAsiaFont,
                    "word/document.xml#paragraph[0]/run[0]", "黑体", "宋体", "east Asian font"),
                new Patch(FontSizeAnalyzer.RuleIdFontSize,
                    "word/document.xml#paragraph[0]/run[0]", "999pt", "12pt", "font size")
            };

            var result = PatchEngine.ApplyPatches(input, patches, output);

            Assert.False(result.Succeeded);
            Assert.Single(result.Failures);
            Assert.False(File.Exists(output));
            Assert.Equal(originalBytes, File.ReadAllBytes(input));
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    [Fact]
    public void PreservesRelationshipsAndUnknownParts()
    {
        var input = CreatePackage("黑体", "Arial", "22", "正文内容");
        var output = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.fixed.docx");
        const string relationships = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"custom/unknown.bin\"/></Relationships>";
        const string unknownPart = "opaque-part-content";
        try
        {
            using (var archive = ZipFile.Open(input, ZipArchiveMode.Update))
            {
                Add(archive, "word/_rels/document.xml.rels", relationships);
                Add(archive, "custom/unknown.bin", unknownPart);
            }

            var before = new FontSizeAnalyzer().Analyze(new WordFormattingReader().Read(input), DemoRuleSet.NkuDraftSample);
            var result = PatchEngine.ApplyPatches(input, PatchEngine.PlanPatches(before), output);

            Assert.True(result.Succeeded);
            using var repaired = ZipFile.OpenRead(output);
            Assert.Equal(relationships, ReadEntry(repaired, "word/_rels/document.xml.rels"));
            Assert.Equal(unknownPart, ReadEntry(repaired, "custom/unknown.bin"));
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    private static string CreateInheritedPackage()
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
        Add(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        Add(archive, "word/document.xml", $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Emphasis\"/><w:highlight w:val=\"yellow\"/></w:rPr><w:t>正文内容</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
        Add(archive, "word/styles.xml", $"<w:styles xmlns:w=\"{W}\"><w:style w:type=\"paragraph\" w:styleId=\"Body\"><w:rPr><w:rFonts w:eastAsia=\"黑体\"/><w:sz w:val=\"22\"/></w:rPr></w:style><w:style w:type=\"character\" w:styleId=\"Emphasis\"><w:rPr><w:b/></w:rPr></w:style></w:styles>");
        return path;
    }

    private static string CreatePackage(string eastAsia, string ascii, string size, string text)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
        Add(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        Add(archive, "word/document.xml", $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:rPr><w:rFonts w:eastAsia=\"{eastAsia}\" w:ascii=\"{ascii}\"/><w:sz w:val=\"{size}\"/></w:rPr><w:t>{text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
        return path;
    }

    private static void Add(ZipArchive archive, string name, string xml)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(xml);
    }

    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry!.Open());
        return reader.ReadToEnd();
    }
}
