using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class ThesisStructurePipelineTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void CompleteSyntheticThesisWithoutHeadingStylesKeepsFrontBodyAndBackMatterSeparate()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln"))) root = root.Parent;
        Assert.NotNull(root);
        var originalPath = Path.Combine(root.FullName, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_complete_thesis.docx");
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var formatted = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".formatted.docx");
        try
        {
            using (var source = ZipFile.OpenRead(originalPath))
            using (var destination = ZipFile.Open(input, ZipArchiveMode.Create))
            {
                foreach (var entry in source.Entries)
                {
                    using var read = entry.Open();
                    using var write = destination.CreateEntry(entry.FullName).Open();
                    if (entry.FullName == "word/document.xml")
                    {
                        var xml = XDocument.Load(read);
                        foreach (var properties in xml.Descendants(W + "pPr").ToArray())
                            properties.Elements().Where(child => child.Name != W + "sectPr" && child.Name != W + "pageBreakBefore").Remove();
                        xml.Descendants(W + "rPr").Remove();
                        xml.Save(write);
                    }
                    else if (entry.FullName == "word/styles.xml")
                    {
                        using var writer = new StreamWriter(write, new UTF8Encoding(false), leaveOpen: true);
                        writer.Write("<w:styles xmlns:w=\"" + W.NamespaceName + "\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia=\"黑体\" w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>");
                    }
                    else read.CopyTo(write);
                }
            }
            var bytes = File.ReadAllBytes(input);
            var snapshot = new WordFormattingReader(inferUnstyledStructure: true, includePackageFingerprint: true).Read(input);
            Assert.True(snapshot.Paragraphs.Count > 150);
            Assert.True(snapshot.Paragraphs.Count(paragraph => paragraph.Scope == ParagraphScopeClassification.ConfirmedBodyProse) > 40);
            Assert.Contains(snapshot.Structure!.Blocks, block => block.Section == ThesisSectionKind.ChineseAbstract);
            Assert.Contains(snapshot.Structure.Blocks, block => block.Section == ThesisSectionKind.EnglishAbstract);
            Assert.Contains(snapshot.Structure.Blocks, block => block.Section == ThesisSectionKind.Contents);
            Assert.Contains(snapshot.Structure.Blocks, block => block.Section == ThesisSectionKind.References);
            Assert.Contains(snapshot.Structure.Blocks, block => block.Section == ThesisSectionKind.Appendix);
            Assert.Contains(snapshot.Structure.Blocks, block => block.Section == ThesisSectionKind.Acknowledgments);
            Assert.All(snapshot.Paragraphs.Where(paragraph => paragraph.Scope == ParagraphScopeClassification.ConfirmedBodyProse),
                paragraph => Assert.Equal(ThesisSectionKind.Body, paragraph.Structure!.Section));
            Assert.Equal(snapshot.Paragraphs.Count, snapshot.SourceMap!.Blocks.Select(block => block.BlockId).Distinct().Count());
            var schemaBefore = SchemaIssues(input);
            var formatter = new ThesisRoleFormatter();
            var plan = formatter.Plan(snapshot, ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1), authorSelectedDraft: true);
            Assert.True(plan.Operations.Count > 50);
            formatter.Apply(input, plan, formatted);
            Assert.Empty(SchemaIssues(formatted).Except(schemaBefore, StringComparer.Ordinal));
            var reread = new WordFormattingReader(inferUnstyledStructure: true).Read(formatted);
            Assert.Equal(snapshot.SourceMap.Blocks.Select(block => block.ExactText), reread.SourceMap!.Blocks.Select(block => block.ExactText));
            Assert.Equal(bytes, File.ReadAllBytes(input));
        }
        finally { File.Delete(input); File.Delete(formatted); }
    }

    private static string[] SchemaIssues(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(document)
            .Select(error => error.Id + "|" + error.Path?.XPath).ToArray();
    }

    [Fact]
    public void UnstyledClosedBodyFlowsThroughExistingAnalysisRepairAndRecheck()
    {
        var input = CreateDocument("摘要", "这是摘要内容，应使用摘要规范而不是误套正文。", "第一章 绪论",
            "这是没有任何标题样式的正文内容，用于验证章节识别后能够正确应用正文规则。 English text.",
            "参考文献", "[1] An example reference.");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var before = File.ReadAllBytes(input);
            var strict = new WordFormattingReader().Read(input);
            Assert.NotEqual(ParagraphScopeClassification.ConfirmedBodyProse, strict.Paragraphs[3].Scope);
            var snapshot = new WordFormattingReader(inferUnstyledStructure: true, includePackageFingerprint: true).Read(input);
            Assert.Equal(ParagraphScopeClassification.ConfirmedBodyProse, snapshot.Paragraphs[3].Scope);
            Assert.NotEqual(ParagraphScopeClassification.ConfirmedBodyProse, snapshot.Paragraphs[1].Scope);
            Assert.NotNull(snapshot.PackageSha256);
            Assert.Equal(snapshot.Paragraphs.Count, snapshot.SourceMap!.Blocks.Count);
            var analysis = new FontSizeAnalyzer().Analyze(snapshot, DemoRuleSet.NkuDraftSample);
            var plan = PatchPlanner.PlanPatches(analysis);
            Assert.NotEmpty(plan.Patches);
            Assert.All(plan.Patches, patch => Assert.Contains("paragraph[3]/", patch.Location));
            var applied = PatchEngine.ApplyPatches(input, plan, output);
            Assert.True(applied.Succeeded);
            var after = new WordFormattingReader(inferUnstyledStructure: true).Read(output);
            Assert.Empty(PatchPlanner.PlanPatches(new FontSizeAnalyzer().Analyze(after, DemoRuleSet.NkuDraftSample)).Patches);
            Assert.Equal("黑体", after.Paragraphs[1].Runs[0].EastAsiaFont); // Abstract was not body-repaired.
            Assert.Equal(snapshot.SourceMap.Blocks.Select(block => block.ExactText), after.SourceMap!.Blocks.Select(block => block.ExactText));
            Assert.Equal(before, File.ReadAllBytes(input));
        }
        finally { File.Delete(input); File.Delete(output); }
    }

    [Fact]
    public async Task ModelProposalsStayReviewOnlyAndAreCachedWithoutChangingSource()
    {
        var input = CreateDocument("没有标题线索的普通文章内容，需要作者确认它在整篇论文中的角色。");
        try
        {
            var original = File.ReadAllBytes(input);
            var snapshot = new WordFormattingReader(inferUnstyledStructure: true).Read(input);
            var model = new EchoModel();
            var coordinator = new ThesisStructureAiCoordinator(model);
            var report = await coordinator.ProposeAsync(snapshot);
            Assert.Single(report.Proposals);
            Assert.True(report.Proposals[0].RequiresReview);
            Assert.Equal(ParagraphScopeClassification.Unknown, snapshot.Paragraphs[0].Scope);
            Assert.Empty(PatchPlanner.PlanPatches(new FontSizeAnalyzer().Analyze(snapshot, DemoRuleSet.NkuDraftSample)).Patches);
            var second = await coordinator.ProposeAsync(snapshot);
            Assert.Equal(0, second.ModelRequests);
            Assert.Equal(1, model.Calls);
            Assert.Equal(original, File.ReadAllBytes(input));

            var approval = new ThesisOutlineApproval(report.DocumentFingerprint,
                [new ApprovedThesisBlock(snapshot.Paragraphs[0].BlockId!, ThesisSectionKind.Body, ThesisBlockKind.Paragraph)]);
            var confirmed = ThesisOutlineReview.Apply(snapshot, approval);
            Assert.Equal(ParagraphScopeClassification.ConfirmedBodyProse, confirmed.Paragraphs[0].Scope);
            var sample = Assert.Single(ThesisOutlineReview.ExportApprovedSamples(snapshot, approval));
            Assert.Equal("author-confirmed", sample.LabelSource);
            Assert.Equal(report.DocumentFingerprint, sample.DocumentGroup);
            Assert.Throws<InvalidDataException>(() => ThesisOutlineReview.Apply(snapshot, approval with { DocumentFingerprint = "stale" }));
        }
        finally { File.Delete(input); }
    }

    [Fact]
    public void RejectsFabricatedIdsDuplicateIdsStaleSnapshotsAndInventedEvidence()
    {
        var request = Request();
        var good = new ThesisModelPrediction("target", ThesisSectionKind.Body, ThesisBlockKind.Paragraph, 0.99, "source text");
        ThesisStructureAiCoordinator.Validate(request, new("fp", [good]));
        foreach (var invalid in new[]
        {
            new ThesisModelResponse("other", [good]),
            new ThesisModelResponse("fp", [good with { BlockId = "invented" }]),
            new ThesisModelResponse("fp", [good with { BlockId = "context" }]),
            new ThesisModelResponse("fp", [good, good]),
            new ThesisModelResponse("fp", [good with { Evidence = "not in source" }]),
            new ThesisModelResponse("fp", [good with { Confidence = double.NaN }]),
            new ThesisModelResponse("fp", [good with { Kind = (ThesisBlockKind)900 }]),
            new ThesisModelResponse("fp", [good with { Kind = ThesisBlockKind.Heading, HeadingLevel = null }]),
            new ThesisModelResponse("fp", [good with { Kind = ThesisBlockKind.TableCell }])
        }) Assert.Throws<InvalidDataException>(() => ThesisStructureAiCoordinator.Validate(request, invalid));
    }

    [Theory]
    [InlineData("{\"documentFingerprint\":\"fp\",\"predictions\":[],\"format\":\"delete all\"}")]
    [InlineData("{\"documentFingerprint\":\"fp\",\"documentFingerprint\":\"fp\",\"predictions\":[]}")]
    [InlineData("{\"documentFingerprint\":\"fp\",\"predictions\":[{\"blockId\":\"target\",\"section\":1,\"kind\":\"Paragraph\",\"confidence\":1,\"evidence\":\"text\"}]}")]
    public void StrictModelJsonCannotInjectPropertiesOrNumericRoles(string json) =>
        Assert.ThrowsAny<Exception>(() => ThesisStructureJson.ParseResponse(json));

    [Theory]
    [InlineData("https://example.com/api/generate")]
    [InlineData("http://user:password@localhost:11434/api/generate")]
    [InlineData("http://localhost:11434/api/generate?redirect=remote")]
    public void OllamaAdapterRejectsNonLocalOrUnexpectedEndpoints(string uri) =>
        Assert.Throws<ArgumentException>(() => new OllamaThesisStructureModel("configured-model", new Uri(uri), localOnlyServiceConfirmed: true));

    [Fact]
    public void OllamaRequiresLocalOnlyServiceAcknowledgmentAndRejectsCloudModels()
    {
        Assert.Throws<ArgumentException>(() => new OllamaThesisStructureModel("configured-model"));
        Assert.Throws<ArgumentException>(() => new OllamaThesisStructureModel("model:cloud", localOnlyServiceConfirmed: true));
    }

    [Fact]
    public async Task OllamaUsesBoundedStructuredOutputWithoutActualNetwork()
    {
        var handler = new RecordingHandler();
        using var model = new OllamaThesisStructureModel("configured-local-model", transport: handler, localOnlyServiceConfirmed: true);
        var response = await model.ClassifyAsync(Request(), CancellationToken.None);
        Assert.Single(response.Predictions);
        using var body = JsonDocument.Parse(handler.Body!);
        Assert.False(body.RootElement.GetProperty("stream").GetBoolean());
        Assert.Equal(JsonValueKind.Object, body.RootElement.GetProperty("format").ValueKind);
        Assert.Equal(0, body.RootElement.GetProperty("options").GetProperty("temperature").GetInt32());
        Assert.Equal("http://127.0.0.1:11434/api/generate", handler.Endpoint);
    }

    [Fact]
    public async Task ModelDeadlineAlsoCancelsAStalledResponseBody()
    {
        using var model = new OllamaThesisStructureModel("configured-local-model", transport: new StallingHandler(),
            localOnlyServiceConfirmed: true, requestTimeout: TimeSpan.FromMilliseconds(150));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => model.ClassifyAsync(Request(), CancellationToken.None));
    }

    [Fact]
    public async Task CallerCancellationIsNotReportedAsSuccessfulAiFallback()
    {
        var input = CreateDocument("未分节的文章正文，需要确认所属部分。");
        try
        {
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            var snapshot = new WordFormattingReader().Read(input);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                new ThesisStructureAiCoordinator(new EchoModel()).ProposeAsync(snapshot, cancellation.Token));
        }
        finally { File.Delete(input); }
    }

    private static ThesisModelRequest Request() => new("fp",
        [new("target", "source text", ThesisSectionKind.Unknown, ThesisBlockKind.Unknown, true),
         new("context", "neighbor text", ThesisSectionKind.Body, ThesisBlockKind.Paragraph, false)]);

    private sealed class EchoModel : IThesisStructureModel
    {
        public int Calls { get; private set; }
        public Task<ThesisModelResponse> ClassifyAsync(ThesisModelRequest request, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Calls++;
            return Task.FromResult(new ThesisModelResponse(request.DocumentFingerprint, request.Blocks.Where(block => block.IsTarget)
                .Select(block => new ThesisModelPrediction(block.BlockId, ThesisSectionKind.Body, ThesisBlockKind.Paragraph, 0.95, block.Text)).ToArray()));
        }
    }

    private sealed class RecordingHandler : HttpMessageHandler
    {
        public string? Body { get; private set; }
        public string? Endpoint { get; private set; }
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Endpoint = request.RequestUri!.AbsoluteUri;
            Body = await request.Content!.ReadAsStringAsync(cancellationToken);
            var prediction = new ThesisModelResponse("fp", [new("target", ThesisSectionKind.Body, ThesisBlockKind.Paragraph, 0.9, "source")]);
            return new HttpResponseMessage(HttpStatusCode.OK)
            { Content = new StringContent(JsonSerializer.Serialize(new { done = true, response = JsonSerializer.Serialize(prediction, ThesisStructureJson.Options) })) };
        }
    }

    private sealed class StallingHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StreamContent(new StallingStream()) });
    }

    private sealed class StallingStream : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        { await Task.Delay(Timeout.Infinite, cancellationToken); return 0; }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private static string CreateDocument(params string[] text)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
        Add("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
        Add("word/document.xml", new XElement(W + "document", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
            new XElement(W + "body", text.Select(value => new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", value)))))).ToString());
        Add("word/styles.xml", "<w:styles xmlns:w=\"" + W.NamespaceName + "\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia=\"黑体\" w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>");
        return path;
        void Add(string name, string content) { using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false)); writer.Write(content); }
    }
}
