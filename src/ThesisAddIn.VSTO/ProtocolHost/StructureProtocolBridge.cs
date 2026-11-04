using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using ThesisAddIn.Protocol;
using ThesisEngine.Core;

/// <summary>Read-only bridge: identifies source paragraphs and emits canonical targets for the Word COM writer.</summary>
internal static class StructureProtocolBridge
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string MainPart = "word/document.xml";

    public static EngineResponse Handle(EngineRequest request, string path)
    {
        if (request.OutputPath is not null || request.RuleId is not null || request.Location is not null || request.Locations?.Count > 0)
            throw new ArgumentException("结构确认/计划/核验是只读操作，不接受输出文件或旧诊断过滤参数。");
        var response = new EngineResponse { RequestId = request.RequestId, Ok = true };
        switch (request.Operation)
        {
            case EngineProtocol.StructureReviewOperation:
                if (request.StructurePlan is not null || request.StructureVerification is not null)
                    throw new ArgumentException("结构读取不接受计划或核验载荷。");
                response.StructureReview = Review(path);
                break;
            case EngineProtocol.StructurePlanOperation:
                if (request.StructurePlan is null || request.StructureVerification is not null)
                    throw new ArgumentException("结构计划需要唯一的 structurePlan 载荷。");
                response.StructurePlan = Plan(path, request.StructurePlan);
                break;
            case EngineProtocol.StructureVerifyOperation:
                if (request.StructureVerification?.PreviousPlan is null || request.StructurePlan is not null)
                    throw new ArgumentException("结构核验需要唯一的 previousPlan。");
                response.StructureVerification = Verify(path, request.StructureVerification.PreviousPlan);
                break;
            default:
                throw new ArgumentException("未知的结构操作。");
        }
        return response;
    }

    private static DocumentSnapshot Read(string path) =>
        new WordFormattingReader(inferUnstyledStructure: true, includePackageFingerprint: true).Read(path);

    private static EngineStructureReviewResult Review(string path)
    {
        var snapshot = Read(path);
        var package = ReadPackage(path, snapshot.PackageSha256!);
        var geometry = Geometry.Create(package.Document, snapshot);
        var decisions = snapshot.Structure!.Blocks.ToDictionary(block => block.BlockId, StringComparer.Ordinal);
        var result = new EngineStructureReviewResult
        {
            DocumentFingerprint = ThesisStructureAiCoordinator.Fingerprint(snapshot), PackageSha256 = snapshot.PackageSha256!,
            SourceParagraphCount = geometry.WordIndices.Length, TotalTableRows = geometry.TotalRows,
            ExpectedWordParagraphCount = geometry.WordIndices.Length + geometry.TotalRows,
            Warnings = snapshot.UnsupportedScopes.Concat(snapshot.Structure.Warnings).Distinct(StringComparer.Ordinal).ToList()
        };
        foreach (var source in snapshot.SourceMap!.Blocks)
        {
            var paragraph = snapshot.Paragraphs[source.ParagraphIndex];
            var decision = decisions[source.BlockId];
            var block = Fill(new EngineStructureBlock(), source, paragraph, geometry);
            block.Section = ToProtocol(decision.Section); block.Kind = ToProtocol(decision.Kind);
            block.HeadingLevel = decision.HeadingLevel; block.Confidence = decision.Confidence;
            block.RequiresReview = decision.RequiresReview; block.Evidence = decision.Evidence;
            block.CanConfirm = source.HasCompleteTextMapping && !paragraph.HasUnsupportedInlineContent && !paragraph.IsTableCell &&
                !string.IsNullOrWhiteSpace(source.ExactText) && block.RunLocations.Count > 0;
            result.Blocks.Add(block);
        }
        return result;
    }

    private static EngineStructureFormatPlan Plan(string path, EngineStructurePlanRequest request)
    {
        if (!request.AuthorSelectedDraft || !Enum.IsDefined(request.TitleMode) || request.ApprovedBlocks is null)
            throw new ArgumentException("整篇计划需要作者明确选择草案配置、标题模式和确认列表。");
        var snapshot = Read(path);
        var fingerprint = ThesisStructureAiCoordinator.Fingerprint(snapshot);
        if (!string.Equals(request.DocumentFingerprint, fingerprint, StringComparison.Ordinal) ||
            !SameHash(request.PackageSha256, snapshot.PackageSha256))
            throw new InvalidOperationException("原文快照已改变，确认大纲已过期；请重新读取结构。");
        var approved = request.ApprovedBlocks.Select(block => block is null
            ? throw new ArgumentException("确认列表包含空项。")
            : new ApprovedThesisBlock(block.BlockId, ToCore(block.Section), ToCore(block.Kind), block.HeadingLevel)).ToArray();
        snapshot = ThesisOutlineReview.Apply(snapshot, new ThesisOutlineApproval(fingerprint, approved));
        var profile = ThesisRoleFormattingProfile.Nku2026Draft(ToCore(request.TitleMode));
        var planned = new ThesisRoleFormatter().Plan(snapshot, profile, authorSelectedDraft: true);
        var package = ReadPackage(path, snapshot.PackageSha256!);
        var geometry = Geometry.Create(package.Document, snapshot);
        var sourceById = snapshot.SourceMap!.Blocks.ToDictionary(block => block.BlockId, StringComparer.Ordinal);
        var selectedIds = approved.Select(block => block.BlockId).ToHashSet(StringComparer.Ordinal);
        var result = new EngineStructureFormatPlan
        {
            DocumentFingerprint = fingerprint, PackageSha256 = snapshot.PackageSha256!,
            SourceParagraphCount = geometry.WordIndices.Length, TotalTableRows = geometry.TotalRows,
            ExpectedWordParagraphCount = geometry.WordIndices.Length + geometry.TotalRows,
            ProfileId = profile.ProfileId, StandardVersion = profile.StandardVersion, ProfileStatus = EngineThesisProfileStatus.Draft,
            TitleMode = request.TitleMode, ProfileSource = profile.Source, AuthorSelectedDraft = true,
            SourceBlocks = snapshot.SourceMap.Blocks.Select(block => Fill(new EngineStructureSourceBlock(), block,
                snapshot.Paragraphs[block.ParagraphIndex], geometry)).ToList(),
            Omissions = planned.Omissions.Select(item => new EngineStructureOmission { BlockId = item.BlockId, Reason = item.Reason }).ToList(),
            Warnings = snapshot.UnsupportedScopes.Concat(snapshot.Structure!.Warnings).Distinct(StringComparer.Ordinal).ToList()
        };
        foreach (var operation in planned.Operations)
        {
            if (!selectedIds.Contains(operation.BlockId))
            {
                result.Omissions.Add(new EngineStructureOmission { BlockId = operation.BlockId, Reason = "作者本轮未选择；保留当前格式。" });
                continue;
            }
            var source = sourceById[operation.BlockId];
            var target = profile.Targets.Single(candidate => candidate.Section == operation.Section && candidate.Kind == operation.Kind &&
                candidate.HeadingLevel == operation.HeadingLevel);
            var transfer = Fill(new EngineStructureFormatOperation(), source, snapshot.Paragraphs[source.ParagraphIndex], geometry);
            transfer.Section = ToProtocol(operation.Section); transfer.Kind = ToProtocol(operation.Kind);
            transfer.HeadingLevel = operation.HeadingLevel; transfer.Confidence = operation.Confidence;
            transfer.RequiresReview = operation.RequiresReview; transfer.Target = ToProtocol(target);
            result.Operations.Add(transfer);
        }
        ValidatePreviousPlan(result);
        return result;
    }

    private static EngineStructureVerificationResult Verify(string path, EngineStructureFormatPlan previous)
    {
        var profile = ValidatePreviousPlan(previous);
        var snapshot = Read(path);
        var package = ReadPackage(path, snapshot.PackageSha256!);
        var geometry = Geometry.Create(package.Document, snapshot);
        var result = new EngineStructureVerificationResult
        {
            TextUnchanged = true, TargetsMatch = true,
            Omissions = previous.Omissions.Select(item => new EngineStructureOmission { BlockId = item.BlockId, Reason = item.Reason }).ToList(),
            Warnings = snapshot.UnsupportedScopes.Distinct(StringComparer.Ordinal).ToList()
        };
        if (previous.SourceParagraphCount != snapshot.SourceMap!.Blocks.Count || previous.TotalTableRows != geometry.TotalRows ||
            previous.ExpectedWordParagraphCount != geometry.WordIndices.Length + geometry.TotalRows)
            AddIssue(result, null, null, "documentStructure", previous.ExpectedWordParagraphCount.ToString(CultureInfo.InvariantCulture),
                (geometry.WordIndices.Length + geometry.TotalRows).ToString(CultureInfo.InvariantCulture), "原文段落或表行数量变化。", text: true);
        if (result.TextUnchanged)
        {
            for (var index = 0; index < previous.SourceBlocks.Count; index++)
            {
                var old = previous.SourceBlocks[index]; var current = snapshot.SourceMap.Blocks[index];
                if (old.ParagraphIndex != current.ParagraphIndex || old.WordParagraphIndex != geometry.WordIndices[index] ||
                    old.StructuralPath != current.StructuralPath || old.ExactText != current.ExactText || old.TextHash != current.TextHash)
                    AddIssue(result, old.BlockId, index, "sourceText", old.TextHash, current.TextHash,
                        "全篇原文、结构路径或原生段落索引变化；停止目标核验，不能将其他同文段落替代此处。", text: true);
            }
        }
        if (result.TextUnchanged)
        {
            var styles = new EffectiveExtraProperties(package.Styles);
            foreach (var operation in previous.Operations)
            {
                var paragraph = snapshot.Paragraphs[operation.ParagraphIndex];
                var source = snapshot.SourceMap.Blocks[operation.ParagraphIndex];
                if (!source.HasCompleteTextMapping || paragraph.IsTableCell || paragraph.HasUnsupportedInlineContent)
                {
                    AddIssue(result, operation.BlockId, operation.ParagraphIndex, "sourceMapping", "safe direct text", "unsupported",
                        "目标块出现未支持结构，不能确认格式已正确写入。");
                    continue;
                }
                var target = profile.Targets.Single(candidate => candidate.Section == ToCore(operation.Section) &&
                    candidate.Kind == ToCore(operation.Kind) && candidate.HeadingLevel == operation.HeadingLevel);
                var beforeCount = result.Issues.Count;
                VerifyTarget(result, operation, target, paragraph, geometry.Paragraphs[operation.ParagraphIndex], styles);
                if (result.Issues.Count == beforeCount) result.VerifiedOperations++;
            }
        }
        result.Succeeded = result.TextUnchanged && result.TargetsMatch && result.VerifiedOperations == previous.Operations.Count;
        return result;
    }

    private static ThesisRoleFormattingProfile ValidatePreviousPlan(EngineStructureFormatPlan plan)
    {
        if (plan is null || !plan.AuthorSelectedDraft || !Enum.IsDefined(plan.TitleMode) || plan.ProfileStatus != EngineThesisProfileStatus.Draft ||
            plan.SourcePartName != MainPart || plan.SourceBlocks is null || plan.Operations is null || plan.Omissions is null ||
            plan.SourceParagraphCount < 0 || plan.SourceParagraphCount != plan.SourceBlocks.Count || plan.TotalTableRows < 0 ||
            plan.ExpectedWordParagraphCount != plan.SourceParagraphCount + plan.TotalTableRows || !ValidHash(plan.PackageSha256))
            throw new ArgumentException("整篇计划缺少有效的源锚点、草案选择或段落几何信息。");
        var profile = ThesisRoleFormattingProfile.Nku2026Draft(ToCore(plan.TitleMode));
        if (plan.ProfileId != profile.ProfileId || plan.StandardVersion != profile.StandardVersion || plan.ProfileSource != profile.Source)
            throw new ArgumentException("整篇计划的规范身份或来源不匹配已实现配置。");
        var sources = new Dictionary<string, EngineStructureSourceBlock>(StringComparer.Ordinal);
        var paths = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < plan.SourceBlocks.Count; index++)
        {
            var source = plan.SourceBlocks[index];
            if (source is null || string.IsNullOrWhiteSpace(source.BlockId) || !sources.TryAdd(source.BlockId, source) ||
                source.PartName != MainPart || source.ParagraphIndex != index || source.WordParagraphIndex < index + 1 ||
                source.WordParagraphIndex > plan.ExpectedWordParagraphCount || string.IsNullOrWhiteSpace(source.StructuralPath) ||
                !paths.Add(source.StructuralPath) || source.ExactText is null || source.TextHash != TextHash(source.ExactText) ||
                source.Location != MainPart + "#paragraph[" + index.ToString(CultureInfo.InvariantCulture) + "]" || source.RunLocations is null)
                throw new ArgumentException("整篇计划包含重复、缺失或不一致的源块。");
        }
        var fingerprint = Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(plan.SourceBlocks.Select(block => new
        { block.BlockId, block.PartName, block.ParagraphIndex, block.StructuralPath, block.ExactText }), ThesisStructureJson.Options)));
        if (fingerprint != plan.DocumentFingerprint) throw new ArgumentException("整篇计划的原文指纹不一致。");
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var operation in plan.Operations)
        {
            if (operation is null || !seen.Add(operation.BlockId) || !sources.TryGetValue(operation.BlockId, out var source) ||
                !SameSource(operation, source) || !operation.HasCompleteTextMapping || operation.IsTableCell ||
                operation.RunLocations.Count == 0 || operation.RequiresReview || !double.IsFinite(operation.Confidence) ||
                operation.Confidence < 0.90 || operation.Confidence > 1)
                throw new ArgumentException("整篇操作含未确认、重复、过期或不支持的源锚点。");
            var section = ToCore(operation.Section); var kind = ToCore(operation.Kind);
            var target = profile.Targets.SingleOrDefault(candidate => candidate.Section == section && candidate.Kind == kind &&
                candidate.HeadingLevel == operation.HeadingLevel);
            if (target is null || operation.Target is null || !SameTarget(operation.Target, ToProtocol(target)))
                throw new ArgumentException("整篇操作目标不属于显式规范配置；不接受客户端任意格式值。");
        }
        return profile;
    }

    private static bool SameSource(EngineStructureSourceBlock first, EngineStructureSourceBlock second) =>
        first.PartName == second.PartName && first.ParagraphIndex == second.ParagraphIndex && first.WordParagraphIndex == second.WordParagraphIndex &&
        first.Location == second.Location && first.FirstRunLocation == second.FirstRunLocation && first.StructuralPath == second.StructuralPath &&
        first.ExactText == second.ExactText && first.TextHash == second.TextHash && first.RunLocations is not null &&
        first.RunLocations.SequenceEqual(second.RunLocations, StringComparer.Ordinal) && first.IsTableCell == second.IsTableCell &&
        first.HasCompleteTextMapping == second.HasCompleteTextMapping;

    private static bool SameTarget(EngineStructureFormatTarget first, EngineStructureFormatTarget second) =>
        first.EastAsiaFont == second.EastAsiaFont && first.LatinFont == second.LatinFont && first.SizePt == second.SizePt &&
        first.Bold == second.Bold && first.Alignment == second.Alignment && first.ExactLineSpacingPt == second.ExactLineSpacingPt &&
        first.SingleLineSpacing == second.SingleLineSpacing && first.BeforePt == second.BeforePt && first.AfterPt == second.AfterPt &&
        first.FirstLineChars == second.FirstLineChars && first.LeftIndentChars == second.LeftIndentChars;

    private static T Fill<T>(T transfer, DocumentSourceBlock source, ParagraphSnapshot paragraph, Geometry geometry) where T : EngineStructureSourceBlock
    {
        transfer.BlockId = source.BlockId; transfer.PartName = source.PartName; transfer.ParagraphIndex = source.ParagraphIndex;
        transfer.WordParagraphIndex = geometry.WordIndices[source.ParagraphIndex];
        transfer.Location = MainPart + "#paragraph[" + source.ParagraphIndex.ToString(CultureInfo.InvariantCulture) + "]";
        transfer.RunLocations = source.Runs.Where(run => run.ExactText.Length > 0).Select(run => run.LegacyLocation).ToList();
        transfer.FirstRunLocation = transfer.RunLocations.FirstOrDefault(); transfer.StructuralPath = source.StructuralPath;
        transfer.ExactText = source.ExactText; transfer.TextHash = source.TextHash; transfer.IsTableCell = paragraph.IsTableCell;
        transfer.HasCompleteTextMapping = source.HasCompleteTextMapping;
        return transfer;
    }

    private static EngineStructureFormatTarget ToProtocol(RoleFormatTarget target) => new()
    {
        EastAsiaFont = target.EastAsiaFont, LatinFont = target.LatinFont, SizePt = target.SizePt, Bold = target.Bold,
        Alignment = target.Alignment is null ? null : Enum.Parse<EngineThesisRoleAlignment>(target.Alignment.Value.ToString()),
        ExactLineSpacingPt = target.ExactLineSpacingPt, SingleLineSpacing = target.SingleLineSpacing,
        BeforePt = target.BeforePt, AfterPt = target.AfterPt, FirstLineChars = target.FirstLineChars, LeftIndentChars = target.LeftIndentChars
    };
    private static EngineThesisSection ToProtocol(ThesisSectionKind value) => Enum.Parse<EngineThesisSection>(value.ToString());
    private static EngineThesisBlockKind ToProtocol(ThesisBlockKind value) => Enum.Parse<EngineThesisBlockKind>(value.ToString());
    private static ThesisSectionKind ToCore(EngineThesisSection value) => Enum.IsDefined(value)
        ? Enum.Parse<ThesisSectionKind>(value.ToString()) : throw new ArgumentException("未知章节类型。");
    private static ThesisBlockKind ToCore(EngineThesisBlockKind value) => Enum.IsDefined(value)
        ? Enum.Parse<ThesisBlockKind>(value.ToString()) : throw new ArgumentException("未知文本块类型。");
    private static ThesisTitleMode ToCore(EngineThesisTitleMode value) => Enum.IsDefined(value)
        ? Enum.Parse<ThesisTitleMode>(value.ToString()) : throw new ArgumentException("未知标题模式。");
    private static bool ValidHash(string? hash) => hash is { Length: 64 } && hash.All(Uri.IsHexDigit);
    private static bool SameHash(string? first, string? second) => ValidHash(first) && ValidHash(second) && string.Equals(first, second, StringComparison.OrdinalIgnoreCase);
    private static string TextHash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();

    private sealed record Package(XDocument Document, XDocument? Styles);
    private static Package ReadPackage(string path, string expectedHash)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (!SameHash(Convert.ToHexString(SHA256.HashData(stream)), expectedHash))
            throw new InvalidOperationException("读取期间源快照已改变，拒绝使用不同的段落几何信息。");
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        XDocument? Load(string name, bool required)
        {
            var entries = archive.Entries.Where(entry => entry.FullName == name).ToArray();
            if (entries.Length == 0 && !required) return null;
            if (entries.Length != 1 || entries[0].Length > 16 * 1024 * 1024) throw new InvalidDataException("快照 part 缺失、重复或过大：" + name);
            using var entryStream = entries[0].Open();
            using var xml = XmlReader.Create(entryStream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16 * 1024 * 1024 });
            return XDocument.Load(xml);
        }
        return new Package(Load(MainPart, true)!, Load("word/styles.xml", false));
    }

    private sealed record Geometry(XElement[] Paragraphs, int[] WordIndices, int TotalRows)
    {
        public static Geometry Create(XDocument document, DocumentSnapshot snapshot)
        {
            var body = document.Root?.Element(W + "body") ?? throw new InvalidDataException("快照缺少 Word body。");
            var paragraphs = new List<XElement>(); var indices = new List<int>(); var completedRows = 0; var totalRows = 0;
            var pending = new Stack<(XElement Element, bool Exit)>(); pending.Push((body, false));
            while (pending.TryPop(out var item))
            {
                if (item.Element.Name == W + "txbxContent") continue;
                if (item.Exit) { if (item.Element.Name == W + "tr") completedRows++; continue; }
                if (item.Element.Name == W + "p")
                {
                    indices.Add(paragraphs.Count + completedRows + 1); paragraphs.Add(item.Element); continue;
                }
                if (item.Element.Name == W + "tr") totalRows++;
                pending.Push((item.Element, true));
                foreach (var child in item.Element.Elements().Reverse()) pending.Push((child, false));
            }
            if (paragraphs.Count != snapshot.Paragraphs.Count) throw new InvalidDataException("Reader 与原生段落几何数量不一致。");
            for (var index = 0; index < paragraphs.Count; index++)
                if (string.Concat(paragraphs[index].Descendants(W + "t").Select(text => text.Value)) != snapshot.Paragraphs[index].Text)
                    throw new InvalidDataException("Reader 与原生段落几何原文不一致。");
            return new Geometry(paragraphs.ToArray(), indices.ToArray(), totalRows);
        }
    }

    private static void VerifyTarget(EngineStructureVerificationResult result, EngineStructureFormatOperation operation, RoleFormatTarget target,
        ParagraphSnapshot paragraph, XElement xmlParagraph, EffectiveExtraProperties styles)
    {
        void Check(string property, string expected, string? actual, bool equal) { if (!equal) AddIssue(result, operation.BlockId, operation.ParagraphIndex,
            property, expected, actual ?? "(unknown)", "写后有效格式未匹配规范目标。"); }
        void Number(string property, double? expected, double? actual) { if (expected is { } value) Check(property,
            value.ToString("0.###", CultureInfo.InvariantCulture), actual?.ToString("0.###", CultureInfo.InvariantCulture),
            actual is { } present && double.IsFinite(present) && Math.Abs(value - present) <= 0.01); }
        var runs = xmlParagraph.Elements(W + "r").Where(run => run.Descendants(W + "t").Any()).ToArray();
        for (var index = 0; index < paragraph.Runs.Count; index++)
        {
            var run = paragraph.Runs[index]; if (run.Text.Length == 0) continue;
            if (run.HasUnresolvedStyle) { Check("runStyle", "resolved", null, false); continue; }
            if (target.EastAsiaFont is { } eastAsia && run.Text.EnumerateRunes().Any(rune => rune.Value is >= 0x3400 and <= 0x9FFF or >= 0x20000 and <= 0x2FA1F))
                Check("eastAsiaFont", eastAsia, run.EastAsiaFont, run.EastAsiaFont == eastAsia);
            if (target.LatinFont is { } latin && ContainsLatinLetter(run.Text))
                Check("latinFont", latin, "ascii=" + run.AsciiSlotFont + "; hAnsi=" + run.HAnsiFont, run.AsciiSlotFont == latin && run.HAnsiFont == latin);
            Number("sizePt", target.SizePt, run.FontSizePt);
            if (target.Bold is { } bold)
            {
                var actual = styles.Bold(xmlParagraph, runs[index]);
                Check("bold", bold.ToString(), actual?.ToString(), actual == bold);
            }
        }
        var format = paragraph.ParagraphFormatting;
        if (format is null || format.HasUnresolvedStyle) { Check("paragraphStyle", "resolved", null, false); return; }
        if (target.Alignment is { } alignment)
        {
            var expected = alignment switch { ThesisRoleAlignment.Left => "left", ThesisRoleAlignment.Center => "center", ThesisRoleAlignment.Right => "right", _ => "both" };
            Check("alignment", expected, format.Alignment, format.Alignment == expected);
        }
        if (target.ExactLineSpacingPt is { } line)
        {
            var expected = "exact:" + Math.Round(line * 20).ToString(CultureInfo.InvariantCulture);
            Check("lineSpacing", expected, format.LineSpacing, format.LineSpacing == expected);
        }
        if (target.SingleLineSpacing is true) Check("lineSpacing", "single", format.LineSpacing, format.LineSpacing == "single");
        Number("beforePt", target.BeforePt, format.SpaceBeforePt); Number("afterPt", target.AfterPt, format.SpaceAfterPt);
        if (target.BeforePt is not null && styles.SpacingOverride(xmlParagraph, before: true) is { } beforeOverride)
            Check("beforePt.override", "explicit point spacing", beforeOverride, false);
        if (target.AfterPt is not null && styles.SpacingOverride(xmlParagraph, before: false) is { } afterOverride)
            Check("afterPt.override", "explicit point spacing", afterOverride, false);
        Number("firstLineChars", target.FirstLineChars, styles.CharacterIndent(xmlParagraph, firstLine: true));
        Number("leftIndentChars", target.LeftIndentChars, styles.CharacterIndent(xmlParagraph, firstLine: false));
    }

    private static bool ContainsLatinLetter(string text) => text.EnumerateRunes().Any(rune => Rune.IsLetter(rune) &&
        rune.Value is >= 0x0041 and <= 0x024F or >= 0x1E00 and <= 0x1EFF or >= 0x2C60 and <= 0x2C7F or
            >= 0xA720 and <= 0xA7FF or >= 0xAB30 and <= 0xAB6F);

    private static void AddIssue(EngineStructureVerificationResult result, string? id, int? index, string property,
        string expected, string actual, string reason, bool text = false)
    {
        if (text) result.TextUnchanged = false; else result.TargetsMatch = false;
        result.Issues.Add(new EngineStructureVerificationIssue { BlockId = id, ParagraphIndex = index, Property = property,
            Expected = expected, Actual = actual, Reason = reason });
    }

    /// <summary>Effective properties and spacing overrides not currently represented by the Core reader.</summary>
    private sealed class EffectiveExtraProperties
    {
        private readonly XDocument? _document;
        private readonly Dictionary<string, XElement> _styles;
        private readonly Dictionary<string, XElement[]?> _chains = new(StringComparer.Ordinal);
        private readonly string? _defaultParagraphStyle;
        public EffectiveExtraProperties(XDocument? styles)
        {
            _document = styles;
            _styles = styles?.Root?.Elements(W + "style").Where(style => style.Attribute(W + "styleId") is not null)
                .ToDictionary(style => (string)style.Attribute(W + "styleId")!, StringComparer.Ordinal) ?? new();
            _defaultParagraphStyle = _styles.Values.FirstOrDefault(style => (string?)style.Attribute(W + "type") == "paragraph" &&
                (string?)style.Attribute(W + "default") is "1" or "true" or "on")?.Attribute(W + "styleId")?.Value;
        }
        private XElement[]? Chain(string? id)
        {
            if (id is null) return Array.Empty<XElement>();
            if (_chains.TryGetValue(id, out var cached)) return cached;
            var result = new List<XElement>(); var seen = new HashSet<string>(StringComparer.Ordinal); var current = id;
            while (current is not null)
            {
                if (!seen.Add(current) || !_styles.TryGetValue(current, out var style)) return _chains[id] = null;
                result.Add(style); current = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
            }
            result.Reverse(); return _chains[id] = result.ToArray();
        }
        private XElement[]? ParagraphChain(XElement paragraph) => Chain((string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") ?? _defaultParagraphStyle);
        public bool? Bold(XElement paragraph, XElement run)
        {
            var paragraphStyles = ParagraphChain(paragraph); var characterStyles = Chain((string?)run.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val"));
            if (paragraphStyles is null || characterStyles is null) return null;
            var bold = ReadBoolean(_document?.Root?.Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr")?.Element(W + "b")) ?? false;
            foreach (var style in paragraphStyles.Concat(characterStyles))
            {
                var property = style.Element(W + "rPr")?.Element(W + "b"); if (property is null) continue;
                var toggle = ReadBoolean(property); if (toggle is null) return null; if (toggle.Value) bold = !bold;
            }
            var direct = run.Element(W + "rPr")?.Element(W + "b");
            return direct is null ? bold : ReadBoolean(direct);
        }
        public double? CharacterIndent(XElement paragraph, bool firstLine)
        {
            var attributes = EffectiveAttributes(paragraph, "ind"); if (attributes is null) return null;
            double? Parse(string? value, double divisor) => double.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? number / divisor : null;
            if (firstLine && attributes.TryGetValue("hangingChars", out var hanging) && Parse(hanging, 100) is { } hangingValue && hangingValue != 0) return -hangingValue;
            var chars = firstLine ? attributes.GetValueOrDefault("firstLineChars") : attributes.GetValueOrDefault("startChars") ?? attributes.GetValueOrDefault("leftChars");
            if (chars is not null) return Parse(chars, 100);
            var points = firstLine ? attributes.GetValueOrDefault("firstLine") : attributes.GetValueOrDefault("start") ?? attributes.GetValueOrDefault("left");
            if (firstLine && attributes.TryGetValue("hanging", out var hangingPoints) && Parse(hangingPoints, 20) is { } hp && hp != 0) return null;
            return points is null || Parse(points, 20) == 0 ? 0 : null;
        }
        public string? SpacingOverride(XElement paragraph, bool before)
        {
            var attributes = EffectiveAttributes(paragraph, "spacing"); if (attributes is null) return "unresolved paragraph style";
            var prefix = before ? "before" : "after";
            if (attributes.TryGetValue(prefix + "Autospacing", out var automatic))
            {
                if (automatic is not ("0" or "false" or "off")) return prefix + "Autospacing=" + automatic;
            }
            if (attributes.TryGetValue(prefix + "Lines", out var lines) &&
                (!int.TryParse(lines, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) || value != 0))
                return prefix + "Lines=" + lines;
            return null;
        }
        private Dictionary<string, string>? EffectiveAttributes(XElement paragraph, string property)
        {
            var styles = ParagraphChain(paragraph); if (styles is null) return null;
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
            void Merge(XElement? properties)
            {
                var element = properties?.Element(W + property);
                if (element is not null)
                    foreach (var attribute in element.Attributes().Where(attribute => attribute.Name.Namespace == W))
                        attributes[attribute.Name.LocalName] = attribute.Value;
            }
            Merge(_document?.Root?.Element(W + "docDefaults")?.Element(W + "pPrDefault")?.Element(W + "pPr"));
            foreach (var style in styles) Merge(style.Element(W + "pPr")); Merge(paragraph.Element(W + "pPr"));
            return attributes;
        }
        private static bool? ReadBoolean(XElement? property) => property is null ? null : (string?)property.Attribute(W + "val") switch
        { null or "1" or "true" or "on" => true, "0" or "false" or "off" => false, _ => null };
    }
}
