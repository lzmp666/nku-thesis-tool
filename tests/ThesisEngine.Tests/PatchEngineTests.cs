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
                new Diagnostic(FontSizeAnalyzer.RuleIdAsciiFont, "warning", "word/document.xml#paragraph[0]/run[0]", "Times New Roman", "Arial", "auto", "font"),
                new Diagnostic(FontSizeAnalyzer.RuleIdFontSize, "warning", "word/document.xml#paragraph[0]/run[0]", "12pt", "(unknown)", "manual", "size")
            ],
            "demo-0.1.0", "covered", []);

        var plan = PatchEngine.PlanPatches(result);

        Assert.Equal(2, plan.Patches.Count);
        Assert.Contains(plan.Patches, patch => patch.RuleId == FontSizeAnalyzer.RuleIdEastAsiaFont
            && patch.ExpectedBefore == "黑体" && patch.DesiredAfter == "宋体");
        Assert.Contains(plan.Patches, patch => patch.RuleId == FontSizeAnalyzer.RuleIdAsciiFont
            && patch.ExpectedBefore == "Arial" && patch.DesiredAfter == "Times New Roman");
        Assert.Single(plan.DeferredDiagnostics);
    }

    [Fact]
    public void AppliesBatchWithoutChangingTextOrInputAndIsIdempotent()
    {
        var input = CreatePackage("黑体", "Arial", "22", "正文内容 Latin");
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
            var run = after.Paragraphs[1].Runs.Single();
            Assert.Equal("宋体", run.EastAsiaFont);
            Assert.Equal(12, run.FontSizePt);
            Assert.Equal("Times New Roman", run.AsciiSlotFont);
            Assert.Equal("Times New Roman", run.HAnsiFont);
            Assert.Equal("正文内容 Latin", run.Text);
            var second = new FontSizeAnalyzer().Analyze(after, DemoRuleSet.NkuDraftSample);
            Assert.Empty(PatchEngine.PlanPatches(second).Patches);
            Assert.Empty(second.Diagnostics);
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    [Fact]
    public void RepairsEveryTextRunAfterAnnotationsInsertReferenceRuns()
    {
        var input = CreatePackage("黑体", "Arial", "21", "第一段 Latin");
        var annotated = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.annotated.docx");
        var output = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.fixed.docx");
        try
        {
            using (var archive = ZipFile.Open(input, ZipArchiveMode.Update))
            {
                var document = XDocument.Parse(ReadEntry(archive, "word/document.xml"));
                var paragraph = document.Descendants(W + "p").ElementAt(1);
                var first = paragraph.Element(W + "r")!;
                foreach (var text in new[] { "第二段 Latin", "第三段 Latin" })
                {
                    var copy = new XElement(first);
                    copy.Element(W + "t")!.Value = text;
                    paragraph.Add(copy);
                }
                archive.GetEntry("word/document.xml")!.Delete();
                Add(archive, "word/document.xml", document.ToString());
            }
            var analyzer = new FontSizeAnalyzer();
            var analysis = analyzer.Analyze(new WordFormattingReader().Read(input), DemoRuleSet.NkuDraftSample);
            WordAnnotationWriter.WriteAnnotatedCopy(input, annotated, analysis.Diagnostics);
            var originalBytes = File.ReadAllBytes(annotated);
            var plan = PatchEngine.PlanPatches(analyzer.Analyze(new WordFormattingReader().Read(annotated), DemoRuleSet.NkuDraftSample));
            Assert.Equal(9, plan.Patches.Count);

            var result = PatchEngine.ApplyPatches(annotated, plan, output);

            Assert.True(result.Succeeded);
            Assert.Equal(originalBytes, File.ReadAllBytes(annotated));
            var after = new WordFormattingReader().Read(output);
            Assert.Equal(new[] { "第一段 Latin", "第二段 Latin", "第三段 Latin" }, after.Paragraphs[1].Runs.Select(run => run.Text));
            Assert.All(after.Paragraphs[1].Runs, run =>
            {
                Assert.Equal("宋体", run.EastAsiaFont);
                Assert.Equal(12, run.FontSizePt);
                Assert.Equal("Times New Roman", run.AsciiFont);
                Assert.Equal("Times New Roman", run.AsciiSlotFont);
                Assert.Equal("Times New Roman", run.HAnsiFont);
            });
            Assert.Empty(analyzer.Analyze(after, DemoRuleSet.NkuDraftSample).Diagnostics);
            using var beforeZip = ZipFile.OpenRead(annotated);
            using var afterZip = ZipFile.OpenRead(output);
            Assert.Equal(ReadEntry(beforeZip, "word/comments.xml"), ReadEntry(afterZip, "word/comments.xml"));
            var beforeDoc = XDocument.Parse(ReadEntry(beforeZip, "word/document.xml"));
            var afterDoc = XDocument.Parse(ReadEntry(afterZip, "word/document.xml"));
            Assert.Equal(beforeDoc.Descendants(W + "r").Where(run => !run.Descendants(W + "t").Any()).Select(run => run.ToString()),
                afterDoc.Descendants(W + "r").Where(run => !run.Descendants(W + "t").Any()).Select(run => run.ToString()));
        }
        finally
        {
            File.Delete(input);
            File.Delete(annotated);
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
            var run = Assert.Single(after.Paragraphs[1].Runs);
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
                "word/document.xml#paragraph[1]/run[0]", "不存在的字体", "宋体", "test");
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
    public void WesternFontRepairSetsAsciiAndHAnsiSlotsAndRemovesThemeOverrides()
    {
        var input = CreatePackageWithFontSlots("宋体", "Arial", "Calibri", "24", "Latin text",
            "w:asciiTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\"");
        var output = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.fixed.docx");
        try
        {
            var analysis = new FontSizeAnalyzer().Analyze(new WordFormattingReader().Read(input), DemoRuleSet.NkuDraftSample);
            var western = Assert.Single(analysis.Diagnostics, diagnostic => diagnostic.RuleId == FontSizeAnalyzer.RuleIdAsciiFont);
            Assert.Equal("ASCII=Arial; HAnsi=Calibri", western.Actual);

            var plan = PatchEngine.PlanPatches(analysis);
            var patch = Assert.Single(plan.Patches, candidate => candidate.RuleId == FontSizeAnalyzer.RuleIdAsciiFont);
            Assert.Equal("ASCII=Arial; HAnsi=Calibri", patch.ExpectedBefore);
            Assert.Equal("Times New Roman", patch.DesiredAfter);

            var applied = PatchEngine.ApplyPatches(input, plan, output);

            Assert.True(applied.Succeeded);
            var repaired = new WordFormattingReader().Read(output);
            var run = Assert.Single(repaired.Paragraphs[1].Runs);
            Assert.Equal("Times New Roman", run.AsciiSlotFont);
            Assert.Equal("Times New Roman", run.HAnsiFont);
            using var archive = ZipFile.OpenRead(output);
            var document = XDocument.Parse(ReadEntry(archive, "word/document.xml"));
            var fonts = Assert.Single(document.Descendants(W + "rFonts"));
            Assert.Equal("Times New Roman", (string?)fonts.Attribute(W + "ascii"));
            Assert.Equal("Times New Roman", (string?)fonts.Attribute(W + "hAnsi"));
            Assert.Null(fonts.Attribute(W + "asciiTheme"));
            Assert.Null(fonts.Attribute(W + "hAnsiTheme"));
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
                    "word/document.xml#paragraph[1]/run[0]", "黑体", "宋体", "east Asian font"),
                new Patch(FontSizeAnalyzer.RuleIdFontSize,
                    "word/document.xml#paragraph[1]/run[0]", "999pt", "12pt", "font size")
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
        Add(archive, "word/document.xml", $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>第一章 绪论</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Emphasis\"/><w:highlight w:val=\"yellow\"/></w:rPr><w:t>正文内容</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
        Add(archive, "word/styles.xml", $"<w:styles xmlns:w=\"{W}\"><w:style w:type=\"paragraph\" w:styleId=\"Body\"><w:rPr><w:rFonts w:eastAsia=\"黑体\"/><w:sz w:val=\"22\"/></w:rPr></w:style><w:style w:type=\"character\" w:styleId=\"Emphasis\"><w:rPr><w:b/></w:rPr></w:style></w:styles>");
        return path;
    }

    private static string CreatePackage(string eastAsia, string ascii, string size, string text)
        => CreatePackageWithFontSlots(eastAsia, ascii, ascii, size, text);

    private static string CreatePackageWithFontSlots(string eastAsia, string ascii, string hAnsi,
        string size, string text, string additionalFontAttributes = "")
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
        Add(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        Add(archive, "word/document.xml", $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>第一章 绪论</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:eastAsia=\"{eastAsia}\" w:ascii=\"{ascii}\" w:hAnsi=\"{hAnsi}\" {additionalFontAttributes}/><w:sz w:val=\"{size}\"/></w:rPr><w:t>{text}</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
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
