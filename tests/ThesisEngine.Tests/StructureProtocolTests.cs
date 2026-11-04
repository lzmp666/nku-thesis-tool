using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using ThesisAddIn.Protocol;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class StructureProtocolTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void TypedReviewPlanAndVerificationRoundTripWithoutCoreTypesInTheContract()
    {
        var request = new EngineRequest
        {
            RequestId = "outline-plan", Operation = EngineProtocol.StructurePlanOperation, InputPath = "test.docx",
            StructurePlan = new EngineStructurePlanRequest
            {
                DocumentFingerprint = "fingerprint", PackageSha256 = new string('A', 64), AuthorSelectedDraft = true,
                TitleMode = EngineThesisTitleMode.Mode2,
                ApprovedBlocks = new() { new() { BlockId = "block-1", Section = EngineThesisSection.Body, Kind = EngineThesisBlockKind.Heading, HeadingLevel = 2 } }
            }
        };
        var json = EngineProtocolCodec.SerializeRequest(request);
        Assert.Contains("\"titleMode\":\"Mode2\"", json, StringComparison.Ordinal);
        Assert.True(EngineProtocolCodec.TryDeserializeRequest(json, out var decoded, out var error), error);
        Assert.Equal(EngineThesisTitleMode.Mode2, decoded!.StructurePlan!.TitleMode);
        Assert.Equal(2, Assert.Single(decoded.StructurePlan.ApprovedBlocks).HeadingLevel);
        Assert.False(EngineProtocolCodec.TryDeserializeRequest(json.Replace("\"Mode2\"", "22", StringComparison.Ordinal), out _, out _));

        var response = new EngineResponse { RequestId = request.RequestId, Ok = true,
            StructureVerification = new EngineStructureVerificationResult { Succeeded = false, TextUnchanged = true, TargetsMatch = false,
                Issues = new() { new() { BlockId = "block-1", Property = "sizePt", Expected = "14", Actual = "11" } } } };
        Assert.True(EngineProtocolCodec.TryDeserializeResponse(EngineProtocolCodec.SerializeResponse(response), out var result, out error), error);
        Assert.False(result!.StructureVerification!.Succeeded);
        Assert.Equal("sizePt", Assert.Single(result.StructureVerification.Issues).Property);
    }

    [Fact]
    public async Task ReviewReturnsNativeParagraphGeometryAndDoesNotOfferUnsafeBlocksForConfirmation()
    {
        using var fixture = new Fixture();
        var table = new XElement(W + "tbl", new XElement(W + "tr", new XElement(W + "tc", P("单元格一"))),
            new XElement(W + "tr", new XElement(W + "tc", P("单元格二"))));
        var tabParagraph = new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "前文"), new XElement(W + "tab"), new XElement(W + "t", "后文")));
        var input = fixture.Create("table.docx", P("第一章 绪论"), P(BodyText), table, P(BodyText + "表后"), tabParagraph, P("参考文献"));
        var before = Hash(input);

        var response = await fixture.Client.ReviewStructureAsync(input);

        Assert.True(response.Ok, response.Error?.Message);
        var review = response.StructureReview!;
        Assert.Equal(7, review.SourceParagraphCount);
        Assert.Equal(2, review.TotalTableRows);
        Assert.Equal(9, review.ExpectedWordParagraphCount);
        Assert.Equal(7, review.Blocks[4].WordParagraphIndex);
        Assert.Equal("word/document.xml#paragraph[4]/run[0]", review.Blocks[4].FirstRunLocation);
        Assert.All(review.Blocks.Where(block => block.IsTableCell), block => Assert.False(block.CanConfirm));
        Assert.False(review.Blocks[5].CanConfirm);
        Assert.False(review.Blocks[5].HasCompleteTextMapping);
        Assert.Equal(before, Hash(input));
    }

    [Fact]
    public async Task PlanOnlyContainsExplicitlySelectedBlocksAndCanonicalTargets()
    {
        using var fixture = new Fixture();
        var input = fixture.CreateStandard();
        var review = await fixture.Review(input);
        var request = Approval(review, review.Blocks[1]);

        var response = await fixture.Client.PlanStructureAsync(input, request);

        Assert.True(response.Ok, response.Error?.Message);
        var plan = response.StructurePlan!;
        var operation = Assert.Single(plan.Operations);
        Assert.Equal(review.Blocks[1].BlockId, operation.BlockId);
        Assert.Equal(12, operation.Target.SizePt);
        Assert.Equal("宋体", operation.Target.EastAsiaFont);
        Assert.Equal("Times New Roman", operation.Target.LatinFont);
        Assert.Equal(EngineThesisRoleAlignment.Justified, operation.Target.Alignment);
        Assert.Equal(20, operation.Target.ExactLineSpacingPt);
        Assert.Equal(2, operation.Target.FirstLineChars);
        Assert.Equal(review.Blocks[1].RunLocations, operation.RunLocations);
        Assert.Equal(review.Blocks.Count, plan.SourceBlocks.Count);
        Assert.Contains(plan.Omissions, omission => omission.BlockId == review.Blocks[2].BlockId && omission.Reason.Contains("作者本轮未选择", StringComparison.Ordinal));
    }

    [Fact]
    public async Task EmptySelectionDoesNotAutomaticallySelectRecognizedBody()
    {
        using var fixture = new Fixture();
        var input = fixture.CreateStandard(); var review = await fixture.Review(input);
        var response = await fixture.Client.PlanStructureAsync(input, Approval(review));

        Assert.True(response.Ok, response.Error?.Message);
        Assert.Empty(response.StructurePlan!.Operations);
        Assert.Contains(response.StructurePlan.Omissions, omission => omission.Reason.Contains("作者本轮未选择", StringComparison.Ordinal));
    }

    [Fact]
    public async Task FormattingOnlyPackageChangeInvalidatesAReviewBeforePlanning()
    {
        using var fixture = new Fixture();
        var input = fixture.CreateStandard(); var review = await fixture.Review(input);
        Fixture.RewriteXml(input, "word/document.xml", document => document.Descendants(W + "p").First().AddFirst(
            new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "center")))));

        var response = await fixture.Client.PlanStructureAsync(input, Approval(review, review.Blocks[1]));

        Assert.False(response.Ok);
        Assert.Equal(EngineProtocol.PatchConflict, response.Error!.Code);
        Assert.Contains("过期", response.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DuplicateApprovalAndReadOnlyOutputPathAreRejected()
    {
        using var fixture = new Fixture();
        var input = fixture.CreateStandard(); var review = await fixture.Review(input);
        var duplicate = Approval(review, review.Blocks[1], review.Blocks[1]);
        var duplicateResult = await fixture.Client.PlanStructureAsync(input, duplicate);
        Assert.False(duplicateResult.Ok);
        var output = fixture.NewPath("should-not-exist.docx");
        var readOnlyResult = await fixture.Client.SendAsync(new EngineRequest
        { Operation = EngineProtocol.StructureReviewOperation, InputPath = input, OutputPath = output });
        Assert.False(readOnlyResult.Ok);
        Assert.Equal(EngineProtocol.InvalidRequest, readOnlyResult.Error!.Code);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task VerifyReportsUnwrittenTargetsAndAcceptsCorrectRoleFormatting()
    {
        using var fixture = new Fixture();
        var input = fixture.CreateStandard(); var review = await fixture.Review(input);
        var request = Approval(review, review.Blocks[0], review.Blocks[1]);
        var response = await fixture.Client.PlanStructureAsync(input, request);
        Assert.True(response.Ok, response.Error?.Message);
        var plan = response.StructurePlan!;
        var unwritten = await fixture.Client.VerifyStructureFormatAsync(input, plan);
        Assert.True(unwritten.Ok, unwritten.Error?.Message);
        Assert.True(unwritten.StructureVerification!.TextUnchanged);
        Assert.False(unwritten.StructureVerification.Succeeded);
        Assert.False(unwritten.StructureVerification.TargetsMatch);

        var formatted = fixture.Format(input, review, request);
        var verified = await fixture.Client.VerifyStructureFormatAsync(formatted, plan);
        Assert.True(verified.Ok, verified.Error?.Message);
        Assert.True(verified.StructureVerification!.Succeeded, string.Join("; ", verified.StructureVerification.Issues.Select(issue => issue.Property + ": " + issue.Actual)));
        Assert.Equal(2, verified.StructureVerification.VerifiedOperations);
        Assert.Empty(verified.StructureVerification.Issues);
    }

    [Fact]
    public async Task VerifyReadsEffectiveFontsFromStylesRatherThanRequiringDirectRunProperties()
    {
        using var fixture = new Fixture();
        var input = fixture.CreateStandard(); var review = await fixture.Review(input); var request = Approval(review, review.Blocks[1]);
        var planResponse = await fixture.Client.PlanStructureAsync(input, request); Assert.True(planResponse.Ok, planResponse.Error?.Message);
        var formatted = fixture.Format(input, review, request);
        Fixture.RewriteXml(formatted, "word/styles.xml", document =>
        {
            var defaults = document.Root!.Element(W + "docDefaults")!.Element(W + "rPrDefault")!.Element(W + "rPr")!;
            defaults.Element(W + "rFonts")!.SetAttributeValue(W + "eastAsia", "宋体");
            defaults.Element(W + "rFonts")!.SetAttributeValue(W + "ascii", "Times New Roman");
            defaults.Element(W + "rFonts")!.SetAttributeValue(W + "hAnsi", "Times New Roman");
            defaults.Element(W + "sz")!.SetAttributeValue(W + "val", "24");
        });
        Fixture.RewriteXml(formatted, "word/document.xml", document =>
            document.Descendants(W + "p").ElementAt(1).Elements(W + "r").Elements(W + "rPr").Remove());

        var response = await fixture.Client.VerifyStructureFormatAsync(formatted, planResponse.StructurePlan!);

        Assert.True(response.Ok, response.Error?.Message);
        Assert.True(response.StructureVerification!.Succeeded, string.Join("; ", response.StructureVerification.Issues.Select(issue => issue.Property + ": " + issue.Actual)));
    }

    [Fact]
    public async Task VerifyProtectsUntouchedTextAndRejectsArbitraryTargetInjection()
    {
        using var fixture = new Fixture();
        var input = fixture.CreateStandard(); var review = await fixture.Review(input); var request = Approval(review, review.Blocks[1]);
        var planned = await fixture.Client.PlanStructureAsync(input, request); Assert.True(planned.Ok, planned.Error?.Message);
        var formatted = fixture.Format(input, review, request);
        Fixture.RewriteXml(formatted, "word/document.xml", document => document.Descendants(W + "p").Last().Descendants(W + "t").First().Value = "作者另改的未选参考文献");
        var changed = await fixture.Client.VerifyStructureFormatAsync(formatted, planned.StructurePlan!);
        Assert.True(changed.Ok, changed.Error?.Message);
        Assert.False(changed.StructureVerification!.TextUnchanged);
        Assert.False(changed.StructureVerification.Succeeded);
        Assert.Equal(0, changed.StructureVerification.VerifiedOperations);

        planned.StructurePlan!.Operations[0].Target.SizePt = 99;
        var injected = await fixture.Client.VerifyStructureFormatAsync(input, planned.StructurePlan);
        Assert.False(injected.Ok);
        Assert.Equal(EngineProtocol.InvalidRequest, injected.Error!.Code);
    }

    [Theory]
    [InlineData("beforeLines", "100")]
    [InlineData("afterLines", "100")]
    [InlineData("beforeAutospacing", "1")]
    [InlineData("afterAutospacing", "on")]
    public async Task VerifyRejectsLineOrAutomaticSpacingThatOverridesTheExpectedPointValue(string property, string value)
    {
        using var fixture = new Fixture();
        var input = fixture.CreateStandard(); var review = await fixture.Review(input); var request = Approval(review, review.Blocks[1]);
        var planned = await fixture.Client.PlanStructureAsync(input, request); Assert.True(planned.Ok, planned.Error?.Message);
        var formatted = fixture.Format(input, review, request);
        Fixture.RewriteXml(formatted, "word/document.xml", document =>
            document.Descendants(W + "p").ElementAt(1).Element(W + "pPr")!.Element(W + "spacing")!.SetAttributeValue(W + property, value));

        var response = await fixture.Client.VerifyStructureFormatAsync(formatted, planned.StructurePlan!);

        Assert.True(response.Ok, response.Error?.Message);
        Assert.True(response.StructureVerification!.TextUnchanged);
        Assert.False(response.StructureVerification.Succeeded);
        Assert.Contains(response.StructureVerification.Issues, issue => issue.Property.EndsWith(".override", StringComparison.Ordinal));
    }

    private const string BodyText = "本研究使用完全合成的普通正文进行结构识别与格式回写验收。This English sentence tests Latin font mapping and does not describe real people.";
    private static XElement P(string text) => new(W + "p", new XElement(W + "r", new XElement(W + "t", text)));
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
    private static EngineStructurePlanRequest Approval(EngineStructureReviewResult review, params EngineStructureBlock[] selected) => new()
    {
        DocumentFingerprint = review.DocumentFingerprint, PackageSha256 = review.PackageSha256,
        AuthorSelectedDraft = true, TitleMode = EngineThesisTitleMode.Mode2,
        ApprovedBlocks = selected.Select(block => new EngineApprovedStructureBlock
        { BlockId = block.BlockId, Section = block.Section, Kind = block.Kind, HeadingLevel = block.HeadingLevel }).ToList()
    };

    private sealed class Fixture : IDisposable
    {
        private readonly string _directory = Path.Combine(Path.GetTempPath(), "nku-structure-protocol-" + Guid.NewGuid().ToString("N"));
        private readonly List<string> _files = new();
        public EngineProtocolClient Client { get; }
        public Fixture()
        {
            Directory.CreateDirectory(_directory);
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln"))) directory = directory.Parent;
            Assert.NotNull(directory);
            var configuration = AppContext.BaseDirectory.Contains(Path.DirectorySeparatorChar + "Release" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ? "Release" : "Debug";
            var host = Path.Combine(directory!.FullName, "src", "ThesisAddIn.VSTO", "ProtocolHost", "bin", configuration, "net10.0", "ThesisEngine.ProtocolHost.dll");
            Assert.True(File.Exists(host), "Build the protocol host before protocol integration tests: " + host);
            Client = new EngineProtocolClient(EngineProcessOptions.ForDotnet(host, TimeSpan.FromSeconds(30)));
        }
        public string NewPath(string name) { var path = Path.Combine(_directory, name); _files.Add(path); return path; }
        public string CreateStandard() => Create("standard.docx", P("第一章 绪论"), P(BodyText), P(BodyText + "第二段"), P("参考文献"), P("[1] 完全合成的参考文献。"));
        public string Create(string name, params XElement[] elements)
        {
            var path = NewPath(name);
            using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
            Write(archive, "word/document.xml", new XDocument(new XElement(W + "document", new XAttribute(XNamespace.Xmlns + "w", W), new XElement(W + "body", elements))));
            Write(archive, "word/styles.xml", new XDocument(new XElement(W + "styles", new XAttribute(XNamespace.Xmlns + "w", W),
                new XElement(W + "docDefaults", new XElement(W + "rPrDefault", new XElement(W + "rPr",
                    new XElement(W + "rFonts", new XAttribute(W + "eastAsia", "黑体"), new XAttribute(W + "ascii", "Arial"), new XAttribute(W + "hAnsi", "Arial")),
                    new XElement(W + "sz", new XAttribute(W + "val", "21")))),
                    new XElement(W + "pPrDefault", new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "line", "360"), new XAttribute(W + "lineRule", "auto"))))),
                new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "default", "1"), new XAttribute(W + "styleId", "Normal")))));
            return path;
        }
        public async Task<EngineStructureReviewResult> Review(string input)
        {
            var response = await Client.ReviewStructureAsync(input); Assert.True(response.Ok, response.Error?.Message); return response.StructureReview!;
        }
        public string Format(string input, EngineStructureReviewResult review, EngineStructurePlanRequest request)
        {
            var snapshot = new WordFormattingReader(inferUnstyledStructure: true, includePackageFingerprint: true).Read(input);
            var confirmed = ThesisOutlineReview.Apply(snapshot, new ThesisOutlineApproval(review.DocumentFingerprint,
                request.ApprovedBlocks.Select(block => new ApprovedThesisBlock(block.BlockId,
                    Enum.Parse<ThesisSectionKind>(block.Section.ToString()), Enum.Parse<ThesisBlockKind>(block.Kind.ToString()), block.HeadingLevel)).ToArray()));
            var formatter = new ThesisRoleFormatter(); var plan = formatter.Plan(confirmed, ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode2), true);
            var selected = request.ApprovedBlocks.Select(block => block.BlockId).ToHashSet(StringComparer.Ordinal);
            plan = plan with { Operations = plan.Operations.Where(operation => selected.Contains(operation.BlockId)).ToArray() };
            var output = NewPath("formatted-" + Guid.NewGuid().ToString("N") + ".docx"); formatter.Apply(input, plan, output); return output;
        }
        public static void RewriteXml(string path, string name, Action<XDocument> mutate)
        {
            using var archive = ZipFile.Open(path, ZipArchiveMode.Update); var entry = archive.GetEntry(name)!;
            XDocument document; using (var stream = entry.Open()) document = XDocument.Load(stream);
            mutate(document); entry.Delete(); Write(archive, name, document);
        }
        private static void Write(ZipArchive archive, string name, XDocument document) { using var stream = archive.CreateEntry(name).Open(); document.Save(stream); }
        public void Dispose()
        {
            foreach (var file in _files.Distinct(StringComparer.Ordinal)) if (File.Exists(file)) File.Delete(file);
            if (Directory.Exists(_directory) && Directory.GetFileSystemEntries(_directory).Length == 0) Directory.Delete(_directory);
        }
    }
}
