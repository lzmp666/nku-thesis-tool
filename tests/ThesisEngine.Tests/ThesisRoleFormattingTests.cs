using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class ThesisRoleFormattingTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void AppliesExplicitDraftTargetsToHighConfidenceRolesAndPreservesOtherPackageParts()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "draft.docx");
        var output = Path.Combine(workspace.Path, "formatted.docx");
        var specs = AllRoleSpecs();
        var profile = ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1);
        var seededIndentIndices = ExplicitIndentTargets(specs, profile);
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray(), includePayloadParts: true,
            seedLegacyIndentIndices: seededIndentIndices);
        var inputBefore = SHA256.HashData(File.ReadAllBytes(input));
        var snapshot = ReadSnapshot(input, specs);
        var formatter = new ThesisRoleFormatter();

        var plan = formatter.Plan(snapshot, profile, authorSelectedDraft: true);
        var result = formatter.Apply(input, plan, output);

        Assert.Equal(specs.Count, result.AppliedParagraphs);
        Assert.Equal(inputBefore, SHA256.HashData(File.ReadAllBytes(input)));
        Assert.Equal("formatted.docx", Path.GetFileName(result.OutputPath));
        Assert.Empty(plan.Omissions);
        AssertPackagePayloadsUnchangedExceptMainDocument(input, output);
        AssertAllRoleTargetsReadBack(output, specs, profile);
    }

    [Fact]
    public void FormatsAppendixCaptionsInTheCompleteUnstyledSyntheticThesisAndCaptionReanalysisIsClean()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "synthetic-unstyled.docx");
        var output = Path.Combine(workspace.Path, "synthetic-formatted.docx");
        CreateUnstyledSyntheticThesisCopy(input);
        var originalPackage = File.ReadAllBytes(input);
        var reader = new WordFormattingReader(inferUnstyledStructure: true, includePackageFingerprint: true);
        var snapshot = reader.Read(input);
        var appendixTableCaption = Assert.Single(snapshot.Structure!.Blocks, block =>
            block.Section == ThesisSectionKind.Appendix && block.Kind == ThesisBlockKind.TableCaption);
        Assert.Equal(6, snapshot.Paragraphs.Count(paragraph => paragraph.Caption is not null));

        var formatter = new ThesisRoleFormatter();
        var plan = formatter.Plan(snapshot,
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1), authorSelectedDraft: true);
        Assert.Contains(plan.Operations, operation => operation.BlockId == appendixTableCaption.BlockId);
        Assert.DoesNotContain(plan.Omissions, omission => omission.BlockId == appendixTableCaption.BlockId);
        formatter.Apply(input, plan, output);

        var reanalysis = new FontSizeAnalyzer().Analyze(
            new WordFormattingReader(inferUnstyledStructure: true).Read(output), DemoRuleSet.NkuDraftSample);
        Assert.DoesNotContain(reanalysis.Diagnostics, diagnostic =>
            diagnostic.RuleId.StartsWith("demo/caption-", StringComparison.Ordinal));
        Assert.Equal(originalPackage, File.ReadAllBytes(input));
    }

    [Fact]
    public void Mode2HeadingTargetsClearLegacyParagraphIndentsAndSetOnlyTheSpecifiedTwoCharacterIndent()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "mode2-source.docx");
        var output = Path.Combine(workspace.Path, "mode2-formatted.docx");
        var specs = new[]
        {
            new BlockSpec("第一章 标题", ThesisSectionKind.Body, ThesisBlockKind.Heading, HeadingLevel: 1),
            new BlockSpec("1.1 标题", ThesisSectionKind.Body, ThesisBlockKind.Heading, HeadingLevel: 2),
            new BlockSpec("1.1.1 标题", ThesisSectionKind.Body, ThesisBlockKind.Heading, HeadingLevel: 3),
            new BlockSpec("1.1.1.1 标题", ThesisSectionKind.Body, ThesisBlockKind.Heading, HeadingLevel: 4)
        };
        var profile = ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode2);
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray(),
            seedLegacyIndentIndices: ExplicitIndentTargets(specs, profile));
        var formatter = new ThesisRoleFormatter();
        var plan = formatter.Plan(ReadSnapshot(input, specs), profile, authorSelectedDraft: true);

        formatter.Apply(input, plan, output);

        var document = XDocument.Parse(Encoding.UTF8.GetString(ReadPackageParts(output)["word/document.xml"]));
        var paragraphs = document.Root!.Element(W + "body")!.Elements(W + "p").ToArray();
        for (var index = 0; index < paragraphs.Length; index++)
        {
            var indentation = paragraphs[index].Element(W + "pPr")?.Element(W + "ind");
            Assert.Equal("0", (string?)indentation?.Attribute(W + "firstLineChars"));
            Assert.Equal(index < 2 ? "0" : "200", (string?)indentation?.Attribute(W + "leftChars"));
            Assert.Null(indentation?.Attribute(W + "firstLine"));
            Assert.Null(indentation?.Attribute(W + "left"));
            Assert.Null(indentation?.Attribute(W + "hanging"));
            Assert.Null(indentation?.Attribute(W + "hangingChars"));
        }
    }

    [Fact]
    public void RejectsPlanWhenTheInputPackageHasChanged()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "source.docx");
        var output = Path.Combine(workspace.Path, "stale-result.docx");
        var specs = new[] { new BlockSpec("正文段落内容用于测试。", ThesisSectionKind.Body, ThesisBlockKind.Paragraph) };
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray());
        var snapshot = ReadSnapshot(input, specs);
        var plan = new ThesisRoleFormatter().Plan(snapshot,
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode2), authorSelectedDraft: true);

        RewriteParagraphText(input, "正文段落内容用于测试。", "原文已修改，计划必须过期。");

        Assert.Throws<InvalidOperationException>(() => new ThesisRoleFormatter().Apply(input, plan, output));
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void RejectsTamperedSourceAnchorEvenWhenPackageFingerprintStillMatches()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "source.docx");
        var output = Path.Combine(workspace.Path, "tampered-result.docx");
        var specs = new[] { new BlockSpec("正文锚点保持精确。", ThesisSectionKind.Body, ThesisBlockKind.Paragraph) };
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray());
        var formatter = new ThesisRoleFormatter();
        var plan = formatter.Plan(ReadSnapshot(input, specs),
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1), authorSelectedDraft: true);
        var operation = Assert.Single(plan.Operations);
        var tamperedPlan = plan with
        {
            Operations = Array.AsReadOnly(new[] { operation with { ExactText = "伪造锚点" } })
        };

        Assert.Throws<InvalidOperationException>(() => formatter.Apply(input, tamperedPlan, output));
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void LeavesLowConfidenceOrReviewRequiredRoleUnchangedInPlan()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "uncertain.docx");
        var specs = new[]
        {
            new BlockSpec("这个段落还不能确定是否属于正文。", ThesisSectionKind.Body,
                ThesisBlockKind.Paragraph, Confidence: 0.89, RequiresReview: false)
        };
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray());
        var snapshot = ReadSnapshot(input, specs);

        var plan = new ThesisRoleFormatter().Plan(snapshot,
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1), authorSelectedDraft: true);

        Assert.Empty(plan.Operations);
        Assert.Contains(plan.Omissions, omission => omission.Reason.Contains("不足", StringComparison.Ordinal));
    }

    [Fact]
    public void DoesNotTreatUnknownTextAsBodyFormattingTarget()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "unknown.docx");
        var spec = new BlockSpec("没有结构边界的长段文本。", ThesisSectionKind.Unknown, ThesisBlockKind.Unknown,
            Confidence: 0.99, RequiresReview: false);
        CreateDocx(input, new[] { spec.Text });
        var snapshot = ReadSnapshot(input, new[] { spec });

        var plan = new ThesisRoleFormatter().Plan(snapshot,
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1), authorSelectedDraft: true);

        Assert.Empty(plan.Operations);
        Assert.Single(plan.Omissions);
        Assert.Contains("未覆盖结构角色", plan.Omissions[0].Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void DoesNotFormatATableCellEvenIfAProposedDecisionCallsItBodyProse()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "table-cell.docx");
        var specs = new[] { new BlockSpec("表格里的文字。", ThesisSectionKind.Body, ThesisBlockKind.Paragraph) };
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray(), tableCellIndices: new HashSet<int> { 0 });

        var plan = new ThesisRoleFormatter().Plan(ReadSnapshot(input, specs),
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1), authorSelectedDraft: true);

        Assert.Empty(plan.Operations);
        Assert.Contains(plan.Omissions, omission => omission.Reason.Contains("表格单元格", StringComparison.Ordinal));
    }

    [Fact]
    public void DoesNotFormatParagraphWithUnsupportedInlineContent()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "unsupported-inline.docx");
        var specs = new[] { new BlockSpec("带制表符的正文。", ThesisSectionKind.Body, ThesisBlockKind.Paragraph) };
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray(), unsupportedInlineIndices: new HashSet<int> { 0 });

        var plan = new ThesisRoleFormatter().Plan(ReadSnapshot(input, specs),
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1), authorSelectedDraft: true);

        Assert.Empty(plan.Operations);
        Assert.Contains(plan.Omissions, omission => omission.Reason.Contains("未支持内联结构", StringComparison.Ordinal));
    }

    [Fact]
    public void RequiresExplicitAuthorSelectionOfDraftProfile()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "draft-gate.docx");
        var specs = new[] { new BlockSpec("正文。", ThesisSectionKind.Body, ThesisBlockKind.Paragraph) };
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray());
        var snapshot = ReadSnapshot(input, specs);

        Assert.Throws<InvalidOperationException>(() => new ThesisRoleFormatter().Plan(snapshot,
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1)));
    }

    [Fact]
    public void RejectsForgedOrInvalidFormattingProfileValues()
    {
        var formatter = new ThesisRoleFormatter();
        var canonical = ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1);
        var invalidSize = canonical.Targets.ToArray();
        invalidSize[0] = invalidSize[0] with { SizePt = 18.25 };
        var profile = canonical with { Targets = Array.AsReadOnly(invalidSize) };

        Assert.Throws<ArgumentOutOfRangeException>(() => formatter.Plan(
            EmptySnapshot(), profile, authorSelectedDraft: true));

        var forged = canonical with { Status = ThesisRoleProfileStatus.Approved };
        Assert.Throws<NotSupportedException>(() => formatter.Plan(
            EmptySnapshot(), forged, authorSelectedDraft: true));
    }

    [Fact]
    public void RefusesToOverwriteSourceOrExistingOutput()
    {
        using var workspace = new TestWorkspace();
        var input = Path.Combine(workspace.Path, "source.docx");
        var existingOutput = Path.Combine(workspace.Path, "existing.docx");
        var specs = new[] { new BlockSpec("正文。", ThesisSectionKind.Body, ThesisBlockKind.Paragraph) };
        CreateDocx(input, specs.Select(spec => spec.Text).ToArray());
        CreateDocx(existingOutput, new[] { "既有文件。" });
        var formatter = new ThesisRoleFormatter();
        var plan = formatter.Plan(ReadSnapshot(input, specs),
            ThesisRoleFormattingProfile.Nku2026Draft(ThesisTitleMode.Mode1), authorSelectedDraft: true);

        Assert.Throws<InvalidOperationException>(() => formatter.Apply(input, plan, input));
        Assert.Throws<IOException>(() => formatter.Apply(input, plan, existingOutput));
        Assert.Equal("existing.docx", Path.GetFileName(existingOutput));
    }

    private static List<BlockSpec> AllRoleSpecs() => new()
    {
        new("中文摘要", ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Heading, HeadingLevel: 1),
        new("本文为中文摘要正文。", ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Paragraph),
        new("Abstract", ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Heading, HeadingLevel: 1),
        new("This is an English abstract paragraph.", ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Paragraph),
        new("目录", ThesisSectionKind.Contents, ThesisBlockKind.Heading, HeadingLevel: 1),
        new("第一章 绪论", ThesisSectionKind.Body, ThesisBlockKind.Heading, HeadingLevel: 1),
        new("1.1 研究问题", ThesisSectionKind.Body, ThesisBlockKind.Heading, HeadingLevel: 2),
        new("1.1.1 调查对象", ThesisSectionKind.Body, ThesisBlockKind.Heading, HeadingLevel: 3),
        new("1.1.1.1 样本说明", ThesisSectionKind.Body, ThesisBlockKind.Heading, HeadingLevel: 4),
        new("正文中英文 mixed text。", ThesisSectionKind.Body, ThesisBlockKind.Paragraph),
        new("图 1.1 图示标题", ThesisSectionKind.Body, ThesisBlockKind.FigureCaption),
        new("表 1.1 表格标题", ThesisSectionKind.Body, ThesisBlockKind.TableCaption),
        new("参考文献", ThesisSectionKind.References, ThesisBlockKind.Heading, HeadingLevel: 1),
        new("参考文献正文。", ThesisSectionKind.References, ThesisBlockKind.Paragraph),
        new("附录", ThesisSectionKind.Appendix, ThesisBlockKind.Heading, HeadingLevel: 1),
        new("附录正文。", ThesisSectionKind.Appendix, ThesisBlockKind.Paragraph),
        new("致谢", ThesisSectionKind.Acknowledgments, ThesisBlockKind.Heading, HeadingLevel: 1),
        new("致谢正文。", ThesisSectionKind.Acknowledgments, ThesisBlockKind.Paragraph),
        new("图A.1　附录图示", ThesisSectionKind.Appendix, ThesisBlockKind.FigureCaption),
        new("表A.1　附录表格", ThesisSectionKind.Appendix, ThesisBlockKind.TableCaption)
    };

    private static void CreateUnstyledSyntheticThesisCopy(string destinationPath)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln"))) root = root.Parent;
        if (root is null) throw new DirectoryNotFoundException("Repository root was not found from the test output directory.");
        var sourcePath = Path.Combine(root.FullName, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_complete_thesis.docx");

        using var source = ZipFile.OpenRead(sourcePath);
        using var destination = ZipFile.Open(destinationPath, ZipArchiveMode.Create);
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

    private static DocumentSnapshot ReadSnapshot(string path, IReadOnlyList<BlockSpec> specs)
    {
        var package = File.ReadAllBytes(path);
        using var stream = new MemoryStream(package, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = archive.GetEntry("word/document.xml") ?? throw new InvalidDataException("No main document part.");
        using var partStream = entry.Open();
        var document = XDocument.Load(partStream, LoadOptions.PreserveWhitespace);
        var body = document.Root?.Element(W + "body") ?? throw new InvalidDataException("No body.");
        var sourceParagraphs = body.Descendants(W + "p").Where(paragraph => !paragraph.Ancestors(W + "txbxContent").Any()).ToArray();
        var sourceMap = DocumentSourceMap.Create("word/document.xml", sourceParagraphs);
        Assert.Equal(specs.Count, sourceMap.Blocks.Count);

        var paragraphs = sourceParagraphs.Select((paragraph, index) => new ParagraphSnapshot(index,
            Array.Empty<TextRunSnapshot>())
        {
            Text = string.Concat(paragraph.Descendants(W + "t").Select(text => text.Value)),
            BlockId = sourceMap.Blocks[index].BlockId,
            IsTableCell = paragraph.Ancestors(W + "tc").Any(),
            HasUnsupportedInlineContent = paragraph.Descendants().Any(element =>
                element.Name == W + "drawing" || element.Name == W + "hyperlink" || element.Name == W + "tab")
        }).ToArray();
        var decisions = specs.Select((spec, index) => new ThesisStructureDecision(
            sourceMap.Blocks[index].BlockId, index, spec.Section, spec.Kind, spec.Confidence,
            "由测试明确提供的结构证据", spec.RequiresReview, spec.HeadingLevel)).ToArray();

        return new DocumentSnapshot("word/document.xml", paragraphs, Array.Empty<string>())
        {
            SourceMap = sourceMap,
            Structure = new ThesisStructureResult(decisions, Array.Empty<string>()),
            PackageSha256 = Convert.ToHexString(SHA256.HashData(package))
        };
    }

    private static DocumentSnapshot EmptySnapshot() => new("word/document.xml",
        Array.Empty<ParagraphSnapshot>(), Array.Empty<string>())
    {
        SourceMap = DocumentSourceMap.Create("word/document.xml", Array.Empty<XElement>()),
        Structure = new ThesisStructureResult(Array.Empty<ThesisStructureDecision>(), Array.Empty<string>()),
        PackageSha256 = new string('A', 64)
    };

    private static void CreateDocx(string path, IReadOnlyList<string> texts, bool includePayloadParts = false,
        IReadOnlySet<int>? tableCellIndices = null, IReadOnlySet<int>? unsupportedInlineIndices = null,
        IReadOnlySet<int>? seedLegacyIndentIndices = null)
    {
        XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var body = new XElement(W + "body");
        for (var index = 0; index < texts.Count; index++)
        {
            var run = new XElement(W + "r", new XElement(W + "t", texts[index]));
            if (unsupportedInlineIndices?.Contains(index) == true) run.Add(new XElement(W + "tab"));
            var paragraph = new XElement(W + "p", run);
            if (seedLegacyIndentIndices?.Contains(index) == true)
                paragraph.AddFirst(new XElement(W + "pPr", new XElement(W + "ind",
                    new XAttribute(W + "firstLine", "480"), new XAttribute(W + "left", "360"),
                    new XAttribute(W + "hanging", "80"), new XAttribute(W + "hangingChars", "100"))));
            if (tableCellIndices?.Contains(index) == true)
                body.Add(new XElement(W + "tbl", new XElement(W + "tr", new XElement(W + "tc", paragraph))));
            else
                body.Add(paragraph);
        }
        body.Add(new XElement(W + "sectPr"));
        var document = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(W + "document", new XAttribute(XNamespace.Xmlns + "w", W),
                new XAttribute(XNamespace.Xmlns + "r", rel), body));

        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
        AddEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
        AddEntry(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
        var documentEntry = archive.CreateEntry("word/document.xml", CompressionLevel.Optimal);
        using (var output = documentEntry.Open()) document.Save(output);
        if (includePayloadParts)
        {
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/></Relationships>");
            AddEntry(archive, "word/media/image1.png", "synthetic image payload"u8.ToArray());
            AddEntry(archive, "customXml/item1.xml", "<test>preserve this package part</test>");
        }
    }

    private static void AddEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(content);
    }

    private static void AddEntry(ZipArchive archive, string path, byte[] content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        stream.Write(content);
    }

    private static void RewriteParagraphText(string path, string oldText, string newText)
    {
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var archive = new ZipArchive(input, ZipArchiveMode.Read))
        {
            foreach (var entry in archive.Entries)
            {
                using var part = entry.Open();
                using var buffer = new MemoryStream();
                part.CopyTo(buffer);
                entries.Add(entry.FullName, buffer.ToArray());
            }
        }
        using var xmlStream = new MemoryStream(entries["word/document.xml"], writable: false);
        var xml = XDocument.Load(xmlStream, LoadOptions.PreserveWhitespace);
        var text = xml.Descendants(W + "t").Single(element => element.Value == oldText);
        text.Value = newText;
        entries["word/document.xml"] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        using var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        using var rewritten = new ZipArchive(output, ZipArchiveMode.Create);
        foreach (var (name, bytes) in entries) AddEntry(rewritten, name, bytes);
    }

    private static void AssertPackagePayloadsUnchangedExceptMainDocument(string input, string output)
    {
        var before = ReadPackageParts(input);
        var after = ReadPackageParts(output);
        Assert.Equal(before.Keys.OrderBy(key => key), after.Keys.OrderBy(key => key));
        foreach (var name in before.Keys.Where(name => name != "word/document.xml"))
            Assert.Equal(before[name], after[name]);
    }

    private static HashSet<int> ExplicitIndentTargets(IReadOnlyList<BlockSpec> specs, ThesisRoleFormattingProfile profile) =>
        specs.Select((spec, index) => (spec, index))
            .Where(item => profile.Targets.Single(target => target.Section == item.spec.Section &&
                target.Kind == item.spec.Kind && target.HeadingLevel == item.spec.HeadingLevel) is { } target &&
                (target.FirstLineChars is not null || target.LeftIndentChars is not null))
            .Select(item => item.index)
            .ToHashSet();

    private static Dictionary<string, byte[]> ReadPackageParts(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var result = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in archive.Entries)
        {
            using var part = entry.Open();
            using var buffer = new MemoryStream();
            part.CopyTo(buffer);
            result.Add(entry.FullName, buffer.ToArray());
        }
        return result;
    }

    private static void AssertAllRoleTargetsReadBack(
        string path,
        IReadOnlyList<BlockSpec> specs,
        ThesisRoleFormattingProfile profile)
    {
        var parts = ReadPackageParts(path);
        var document = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
        var paragraphs = document.Root!.Element(W + "body")!.Elements(W + "p").ToArray();
        Assert.Equal(specs.Count, paragraphs.Length);
        for (var index = 0; index < specs.Count; index++)
        {
            var spec = specs[index];
            var target = profile.Targets.Single(candidate => candidate.Section == spec.Section &&
                candidate.Kind == spec.Kind && candidate.HeadingLevel == spec.HeadingLevel);
            var paragraph = paragraphs[index];
            Assert.Equal(spec.Text, string.Concat(paragraph.Descendants(W + "t").Select(text => text.Value)));
            var pPr = paragraph.Element(W + "pPr");
            var run = paragraph.Element(W + "r")!;
            var rPr = run.Element(W + "rPr");
            Assert.Same(pPr, paragraph.Elements().First());
            Assert.Same(rPr, run.Elements().First());
            if (target.EastAsiaFont is not null)
                Assert.Equal(target.EastAsiaFont, (string?)rPr?.Element(W + "rFonts")?.Attribute(W + "eastAsia"));
            if (target.LatinFont is not null)
            {
                Assert.Equal(target.LatinFont, (string?)rPr?.Element(W + "rFonts")?.Attribute(W + "ascii"));
                Assert.Equal(target.LatinFont, (string?)rPr?.Element(W + "rFonts")?.Attribute(W + "hAnsi"));
            }
            if (target.SizePt is { } size)
            {
                Assert.Equal(((int)(size * 2)).ToString(), (string?)rPr?.Element(W + "sz")?.Attribute(W + "val"));
                Assert.Equal(((int)(size * 2)).ToString(), (string?)rPr?.Element(W + "szCs")?.Attribute(W + "val"));
            }
            if (target.Bold is true)
            {
                Assert.NotNull(rPr?.Element(W + "b"));
                Assert.NotNull(rPr?.Element(W + "bCs"));
            }
            if (target.Alignment is { } alignment)
            {
                var expected = alignment switch
                {
                    ThesisRoleAlignment.Left => "left",
                    ThesisRoleAlignment.Center => "center",
                    ThesisRoleAlignment.Right => "right",
                    ThesisRoleAlignment.Justified => "both",
                    _ => throw new ArgumentOutOfRangeException()
                };
                Assert.Equal(expected, (string?)pPr?.Element(W + "jc")?.Attribute(W + "val"));
            }
            if (target.ExactLineSpacingPt is { } exact)
            {
                Assert.Equal(((int)(exact * 20)).ToString(), (string?)pPr?.Element(W + "spacing")?.Attribute(W + "line"));
                Assert.Equal("exact", (string?)pPr?.Element(W + "spacing")?.Attribute(W + "lineRule"));
            }
            if (target.SingleLineSpacing is true)
            {
                Assert.Equal("240", (string?)pPr?.Element(W + "spacing")?.Attribute(W + "line"));
                Assert.Equal("auto", (string?)pPr?.Element(W + "spacing")?.Attribute(W + "lineRule"));
            }
            if (target.BeforePt is { } before)
                Assert.Equal(((int)(before * 20)).ToString(), (string?)pPr?.Element(W + "spacing")?.Attribute(W + "before"));
            if (target.AfterPt is { } after)
                Assert.Equal(((int)(after * 20)).ToString(), (string?)pPr?.Element(W + "spacing")?.Attribute(W + "after"));
            if (target.FirstLineChars is { } firstLine)
            {
                Assert.Equal((firstLine * 100).ToString(), (string?)pPr?.Element(W + "ind")?.Attribute(W + "firstLineChars"));
                Assert.Null(pPr?.Element(W + "ind")?.Attribute(W + "firstLine"));
                Assert.Null(pPr?.Element(W + "ind")?.Attribute(W + "hanging"));
                Assert.Null(pPr?.Element(W + "ind")?.Attribute(W + "hangingChars"));
            }
            if (target.LeftIndentChars is { } left)
            {
                Assert.Equal((left * 100).ToString(), (string?)pPr?.Element(W + "ind")?.Attribute(W + "leftChars"));
                Assert.Null(pPr?.Element(W + "ind")?.Attribute(W + "left"));
            }
        }
    }

    private sealed record BlockSpec(
        string Text,
        ThesisSectionKind Section,
        ThesisBlockKind Kind,
        int? HeadingLevel = null,
        double Confidence = 0.99,
        bool RequiresReview = false);

    private sealed class TestWorkspace : IDisposable
    {
        public TestWorkspace()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "nku-role-format-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
        }
    }
}
