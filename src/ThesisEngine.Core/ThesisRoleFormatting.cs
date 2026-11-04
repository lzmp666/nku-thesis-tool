using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;

namespace ThesisEngine.Core;

public enum ThesisTitleMode
{
    Mode1,
    Mode2
}

public enum ThesisRoleAlignment
{
    Left,
    Center,
    Right,
    Justified
}

public enum ThesisRoleProfileStatus
{
    Draft,
    Approved,
    Custom
}

/// <summary>
/// An explicit, partial direct-formatting override for one structural role.
/// Null properties are preserved from the source document.
/// </summary>
public sealed record RoleFormatTarget(
    ThesisSectionKind Section,
    ThesisBlockKind Kind,
    int? HeadingLevel = null,
    string? EastAsiaFont = null,
    string? LatinFont = null,
    double? SizePt = null,
    bool? Bold = null,
    ThesisRoleAlignment? Alignment = null,
    double? ExactLineSpacingPt = null,
    bool? SingleLineSpacing = null,
    double? BeforePt = null,
    double? AfterPt = null,
    int? FirstLineChars = null,
    int? LeftIndentChars = null);

/// <summary>
/// A named, explicitly selected set of role targets. The NKU profile in this prototype
/// remains a draft and is never represented as an approved university rule set.
/// </summary>
public sealed record ThesisRoleFormattingProfile(
    string ProfileId,
    string StandardVersion,
    ThesisRoleProfileStatus Status,
    ThesisTitleMode TitleMode,
    string Source,
    IReadOnlyList<RoleFormatTarget> Targets)
{
    public const string Nku2026DraftProfileId = "nku-graduate-2026-role-format-draft";
    public const string Nku2026DraftSource = "南开大学研究生院《南开大学研究生学位论文写作规范（2026版）》官方 DOCX（https://graduate.nankai.edu.cn/_upload/article/files/46/28/7965029e47d08d27f560f8393fd8/28100848-9be2-4f1b-8f74-5899900cf936.docx）；2026-10-26 独立核验第 3.3 节及表 4.4–4.7；源 SHA-256：0908DE8017A0F8894FFFF4A4EC54A14334FCFE6839A422FCC44017093A9D6C06。该配置仍为作者显式选择的 Draft 预览，未完整覆盖规范，不代表批准规则集。";

    public static ThesisRoleFormattingProfile Nku2026Draft(ThesisTitleMode mode)
    {
        if (!Enum.IsDefined(mode)) throw new ArgumentOutOfRangeException(nameof(mode));
        var chapterAlignment = mode == ThesisTitleMode.Mode1
            ? ThesisRoleAlignment.Left
            : ThesisRoleAlignment.Center;
        var targets = new List<RoleFormatTarget>
        {
            new(ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Heading, 1,
                EastAsiaFont: "黑体", SizePt: 18, Bold: true, Alignment: ThesisRoleAlignment.Center,
                SingleLineSpacing: true, BeforePt: 24, AfterPt: 18, FirstLineChars: 0, LeftIndentChars: 0),
            new(ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Paragraph,
                EastAsiaFont: "宋体", SizePt: 12, ExactLineSpacingPt: 20, BeforePt: 0, AfterPt: 0),
            new(ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Heading, 1,
                LatinFont: "Arial", SizePt: 18, Bold: true, Alignment: ThesisRoleAlignment.Center,
                SingleLineSpacing: true, BeforePt: 24, AfterPt: 18, FirstLineChars: 0, LeftIndentChars: 0),
            new(ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Paragraph,
                LatinFont: "Times New Roman", SizePt: 12, ExactLineSpacingPt: 20, BeforePt: 0, AfterPt: 0),
            new(ThesisSectionKind.Contents, ThesisBlockKind.Heading, 1,
                EastAsiaFont: "黑体", SizePt: 16, Bold: true, Alignment: ThesisRoleAlignment.Center,
                SingleLineSpacing: true, BeforePt: 24, AfterPt: 18, FirstLineChars: 0, LeftIndentChars: 0),
            new(ThesisSectionKind.Body, ThesisBlockKind.Heading, 1,
                EastAsiaFont: "黑体", SizePt: 16, Bold: true, Alignment: chapterAlignment,
                SingleLineSpacing: true, BeforePt: 24, AfterPt: 18, FirstLineChars: 0, LeftIndentChars: 0),
            new(ThesisSectionKind.Body, ThesisBlockKind.Heading, 2,
                EastAsiaFont: "黑体", SizePt: 14, Bold: true, Alignment: chapterAlignment,
                SingleLineSpacing: true, BeforePt: 24, AfterPt: 6, FirstLineChars: 0, LeftIndentChars: 0),
            new(ThesisSectionKind.Body, ThesisBlockKind.Heading, 3,
                EastAsiaFont: "黑体", SizePt: 13, Alignment: ThesisRoleAlignment.Left,
                SingleLineSpacing: true, BeforePt: 12, AfterPt: 6,
                FirstLineChars: 0, LeftIndentChars: mode == ThesisTitleMode.Mode2 ? 2 : 0),
            new(ThesisSectionKind.Body, ThesisBlockKind.Heading, 4,
                EastAsiaFont: "黑体", SizePt: 12, Alignment: ThesisRoleAlignment.Left,
                SingleLineSpacing: true, BeforePt: 12, AfterPt: 6,
                FirstLineChars: 0, LeftIndentChars: mode == ThesisTitleMode.Mode2 ? 2 : 0),
            new(ThesisSectionKind.Body, ThesisBlockKind.Paragraph,
                EastAsiaFont: "宋体", LatinFont: "Times New Roman", SizePt: 12,
                Alignment: ThesisRoleAlignment.Justified, ExactLineSpacingPt: 20,
                BeforePt: 0, AfterPt: 0, FirstLineChars: 2),
            ChapterTitleTarget(ThesisSectionKind.References, chapterAlignment),
            new(ThesisSectionKind.References, ThesisBlockKind.Paragraph,
                EastAsiaFont: "宋体", LatinFont: "Times New Roman", SizePt: 10.5,
                ExactLineSpacingPt: 16, BeforePt: 0, AfterPt: 0),
            ChapterTitleTarget(ThesisSectionKind.Appendix, chapterAlignment),
            new(ThesisSectionKind.Appendix, ThesisBlockKind.Paragraph,
                EastAsiaFont: "宋体", LatinFont: "Times New Roman", SizePt: 12,
                Alignment: ThesisRoleAlignment.Justified, ExactLineSpacingPt: 20,
                BeforePt: 0, AfterPt: 0, FirstLineChars: 2),
            ChapterTitleTarget(ThesisSectionKind.Acknowledgments, chapterAlignment),
            new(ThesisSectionKind.Acknowledgments, ThesisBlockKind.Paragraph,
                EastAsiaFont: "仿宋", SizePt: 12, ExactLineSpacingPt: 16, BeforePt: 0, AfterPt: 0),
            new(ThesisSectionKind.Body, ThesisBlockKind.FigureCaption,
                EastAsiaFont: "宋体", SizePt: 10.5, Alignment: ThesisRoleAlignment.Center,
                SingleLineSpacing: true, BeforePt: 6, AfterPt: 12),
            new(ThesisSectionKind.Body, ThesisBlockKind.TableCaption,
                EastAsiaFont: "宋体", SizePt: 10.5, Alignment: ThesisRoleAlignment.Center,
                SingleLineSpacing: true, BeforePt: 6, AfterPt: 6),
            new(ThesisSectionKind.Appendix, ThesisBlockKind.FigureCaption,
                EastAsiaFont: "宋体", SizePt: 10.5, Alignment: ThesisRoleAlignment.Center,
                SingleLineSpacing: true, BeforePt: 6, AfterPt: 12),
            new(ThesisSectionKind.Appendix, ThesisBlockKind.TableCaption,
                EastAsiaFont: "宋体", SizePt: 10.5, Alignment: ThesisRoleAlignment.Center,
                SingleLineSpacing: true, BeforePt: 6, AfterPt: 6)
        };

        return new ThesisRoleFormattingProfile(Nku2026DraftProfileId, "2026",
            ThesisRoleProfileStatus.Draft, mode, Nku2026DraftSource, Array.AsReadOnly(targets.ToArray()));
    }

    private static RoleFormatTarget ChapterTitleTarget(ThesisSectionKind section, ThesisRoleAlignment alignment) =>
        new(section, ThesisBlockKind.Heading, 1, EastAsiaFont: "黑体", SizePt: 16, Bold: true,
            Alignment: alignment, SingleLineSpacing: true, BeforePt: 24, AfterPt: 18,
            FirstLineChars: 0, LeftIndentChars: 0);
}

public sealed record ThesisRoleFormatOperation(
    string BlockId,
    int ParagraphIndex,
    string StructuralPath,
    string ExactText,
    ThesisSectionKind Section,
    ThesisBlockKind Kind,
    int? HeadingLevel,
    double Confidence,
    bool RequiresReview);

public sealed record ThesisRoleFormatOmission(string? BlockId, string Reason);

public sealed record ThesisRoleFormatPlan(
    string PackageSha256,
    string SourcePartName,
    string ProfileId,
    string StandardVersion,
    ThesisRoleProfileStatus ProfileStatus,
    ThesisTitleMode TitleMode,
    string ProfileSource,
    bool AuthorSelectedDraft,
    IReadOnlyList<ThesisRoleFormatOperation> Operations,
    IReadOnlyList<ThesisRoleFormatOmission> Omissions);

public sealed record ThesisRoleFormatApplyResult(
    string OutputPath,
    int AppliedParagraphs,
    string OutputSha256,
    IReadOnlyList<string> Warnings);

/// <summary>
/// Plans and applies only explicit role targets to high-confidence, source-mapped DOCX
/// paragraphs. It changes direct paragraph/run properties and preserves all text and
/// non-document package parts. This is a draft preview mechanism, not an approved ruleset.
/// </summary>
public sealed class ThesisRoleFormatter
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string MainDocumentPart = "word/document.xml";
    private const double MinimumAutomaticConfidence = 0.90;
    private const int MaximumPackageBytes = 512 * 1024 * 1024;
    private const int MaximumDocumentXmlBytes = 64 * 1024 * 1024;

    private static readonly IReadOnlyDictionary<string, int> RunPropertyOrder = new Dictionary<string, int>(StringComparer.Ordinal)
    {
        ["rStyle"] = 0, ["rFonts"] = 1, ["b"] = 2, ["bCs"] = 3, ["i"] = 4, ["iCs"] = 5,
        ["caps"] = 6, ["smallCaps"] = 7, ["strike"] = 8, ["dstrike"] = 9, ["outline"] = 10,
        ["shadow"] = 11, ["emboss"] = 12, ["imprint"] = 13, ["noProof"] = 14, ["snapToGrid"] = 15,
        ["vanish"] = 16, ["webHidden"] = 17, ["color"] = 18, ["spacing"] = 19, ["w"] = 20,
        ["kern"] = 21, ["position"] = 22, ["sz"] = 23, ["szCs"] = 24, ["highlight"] = 25,
        ["u"] = 26, ["effect"] = 27, ["bdr"] = 28, ["shd"] = 29, ["fitText"] = 30,
        ["vertAlign"] = 31, ["rtl"] = 32, ["cs"] = 33, ["em"] = 34, ["lang"] = 35,
        ["eastAsianLayout"] = 36, ["specVanish"] = 37, ["oMath"] = 38
    };

    private static readonly IReadOnlyDictionary<string, int> ParagraphPropertyOrder = new Dictionary<string, int>(StringComparer.Ordinal)
    {
        ["pStyle"] = 0, ["keepNext"] = 1, ["keepLines"] = 2, ["pageBreakBefore"] = 3,
        ["framePr"] = 4, ["widowControl"] = 5, ["numPr"] = 6, ["suppressLineNumbers"] = 7,
        ["pBdr"] = 8, ["shd"] = 9, ["tabs"] = 10, ["suppressAutoHyphens"] = 11,
        ["kinsoku"] = 12, ["wordWrap"] = 13, ["overflowPunct"] = 14, ["topLinePunct"] = 15,
        ["autoSpaceDE"] = 16, ["autoSpaceDN"] = 17, ["bidi"] = 18, ["adjustRightInd"] = 19,
        ["snapToGrid"] = 20, ["spacing"] = 21, ["ind"] = 22, ["contextualSpacing"] = 23,
        ["mirrorIndents"] = 24, ["suppressOverlap"] = 25, ["jc"] = 26, ["textDirection"] = 27,
        ["textAlignment"] = 28, ["textboxTightWrap"] = 29, ["outlineLvl"] = 30, ["divId"] = 31,
        ["cnfStyle"] = 32, ["rPr"] = 33, ["sectPr"] = 34, ["pPrChange"] = 35
    };

    public ThesisRoleFormatPlan Plan(
        DocumentSnapshot snapshot,
        ThesisRoleFormattingProfile profile,
        bool authorSelectedDraft = false)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(profile);
        ValidateProfile(profile);
        if (!authorSelectedDraft)
            throw new InvalidOperationException("此排版配置仍是 2026 草案，必须由作者显式选择后才能生成计划。");
        if (snapshot.SourceMap is null || snapshot.Structure is null)
            throw new ArgumentException("排版计划需要原文映射和结构分类。", nameof(snapshot));
        if (!string.Equals(snapshot.PartName, MainDocumentPart, StringComparison.Ordinal) ||
            !string.Equals(snapshot.SourceMap.PartName, MainDocumentPart, StringComparison.Ordinal))
            throw new NotSupportedException("此原型只处理 word/document.xml 主文档部分。");
        var packageHash = NormalizeSha256(snapshot.PackageSha256, nameof(snapshot.PackageSha256));

        var sourceById = snapshot.SourceMap.Blocks.ToDictionary(block => block.BlockId, StringComparer.Ordinal);
        var paragraphById = snapshot.Paragraphs
            .Where(paragraph => paragraph.BlockId is not null)
            .GroupBy(paragraph => paragraph.BlockId!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        var decisions = snapshot.Structure.Blocks;
        if (decisions.Select(decision => decision.BlockId).Distinct(StringComparer.Ordinal).Count() != decisions.Count)
            throw new InvalidDataException("结构结果包含重复 BlockId，不能生成可验证的排版计划。");

        var operations = new List<ThesisRoleFormatOperation>();
        var omissions = new List<ThesisRoleFormatOmission>();
        foreach (var decision in decisions)
        {
            var target = profile.Targets.SingleOrDefault(candidate => Matches(candidate, decision));
            if (target is null)
            {
                omissions.Add(new ThesisRoleFormatOmission(decision.BlockId,
                    $"未覆盖结构角色：{decision.Section}/{decision.Kind}/{decision.HeadingLevel?.ToString() ?? "无标题级别"}；Draft profile 未给出可验证目标值，保留原格式。"));
                continue;
            }

            if (!sourceById.TryGetValue(decision.BlockId, out var source))
            {
                omissions.Add(new ThesisRoleFormatOmission(decision.BlockId, "没有与结构块匹配的原文 SourceMap 锚点；保留原格式。"));
                continue;
            }
            if (!paragraphById.TryGetValue(decision.BlockId, out var paragraphMatches) || paragraphMatches.Length != 1)
            {
                omissions.Add(new ThesisRoleFormatOmission(decision.BlockId, "快照段落映射缺失或不唯一；保留原格式。"));
                continue;
            }
            var paragraph = paragraphMatches[0];
            if (!double.IsFinite(decision.Confidence) || decision.Confidence is < 0 or > 1 ||
                decision.Confidence < MinimumAutomaticConfidence || decision.RequiresReview)
            {
                omissions.Add(new ThesisRoleFormatOmission(decision.BlockId,
                    $"结构置信度 {decision.Confidence:0.00} 不足、无效或仍需复核；保留原格式。"));
                continue;
            }
            if (decision.ParagraphIndex != source.ParagraphIndex || paragraph.ParagraphIndex != source.ParagraphIndex ||
                !string.Equals(paragraph.Text, source.ExactText, StringComparison.Ordinal))
            {
                omissions.Add(new ThesisRoleFormatOmission(decision.BlockId, "结构、Reader 段落与 SourceMap 的段号或精确文本不一致；保留原格式。"));
                continue;
            }
            if (!source.HasCompleteTextMapping || source.UnsupportedReasons.Count > 0 ||
                paragraph.HasUnsupportedInlineContent || paragraph.IsTableCell)
            {
                omissions.Add(new ThesisRoleFormatOmission(decision.BlockId, "原文块含未支持内联结构或表格单元格，整段保留原格式。"));
                continue;
            }
            if (string.IsNullOrWhiteSpace(source.ExactText))
            {
                omissions.Add(new ThesisRoleFormatOmission(decision.BlockId, "空段落没有可验证的文本锚点；保留原格式。"));
                continue;
            }

            operations.Add(new ThesisRoleFormatOperation(source.BlockId, source.ParagraphIndex,
                source.StructuralPath, source.ExactText, decision.Section, decision.Kind,
                decision.HeadingLevel, decision.Confidence, decision.RequiresReview));
        }

        return new ThesisRoleFormatPlan(packageHash, snapshot.SourceMap.PartName, profile.ProfileId,
            profile.StandardVersion, profile.Status, profile.TitleMode, profile.Source, authorSelectedDraft,
            Array.AsReadOnly(operations.ToArray()), Array.AsReadOnly(omissions.ToArray()));
    }

    public ThesisRoleFormatApplyResult Apply(string inputPath, ThesisRoleFormatPlan plan, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ValidatePlan(plan);
        if (plan.Operations.Count == 0)
            throw new InvalidOperationException("计划中没有可安全应用的高置信度格式目标。");

        var fullInputPath = Path.GetFullPath(inputPath);
        var fullOutputPath = Path.GetFullPath(outputPath);
        if (!string.Equals(Path.GetExtension(fullInputPath), ".docx", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Path.GetExtension(fullOutputPath), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("输入和新输出必须使用 .docx 扩展名。");
        if (string.Equals(fullInputPath, fullOutputPath, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("排版只能另存为新文件，不能覆盖原文。");
        if (!Directory.Exists(Path.GetDirectoryName(fullOutputPath)!))
            throw new DirectoryNotFoundException("输出目录必须已存在。");
        if (File.Exists(fullOutputPath) || Directory.Exists(fullOutputPath))
            throw new IOException("输出目标已存在；不会覆盖已有文件。");

        var tempPath = Path.Combine(Path.GetDirectoryName(fullOutputPath)!,
            ".nku-role-format-" + Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            using var inputStream = new FileStream(fullInputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (inputStream.Length <= 0 || inputStream.Length > MaximumPackageBytes)
                throw new InvalidDataException("DOCX 包大小不在支持范围内。");
            var actualHash = SHA256.HashData(inputStream);
            if (!CryptographicOperations.FixedTimeEquals(actualHash, Convert.FromHexString(plan.PackageSha256)))
                throw new InvalidOperationException("原文包已变化，排版计划过期；请重新分析文档。");
            inputStream.Position = 0;

            Dictionary<string, PartFingerprint> expectedParts;
            using (var sourceArchive = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: true))
            {
                ValidateArchiveEntries(sourceArchive);
                if (sourceArchive.Entries.Any(entry => entry.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase) ||
                    entry.FullName.StartsWith("word/_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
                    throw new NotSupportedException("DOCX 包含数字签名；修改 XML 会使签名失效，因此拒绝生成输出。");
                var documentEntry = FindUniqueEntry(sourceArchive, MainDocumentPart);
                if (documentEntry.Length > MaximumDocumentXmlBytes)
                    throw new InvalidDataException("主文档 XML 超过原型支持大小。");
                var document = LoadXml(documentEntry);
                var body = document.Root?.Element(W + "body")
                    ?? throw new InvalidDataException("DOCX 缺少主文档 body。");
                var sourceParagraphs = GetMainStoryParagraphs(body);
                var currentMap = DocumentSourceMap.Create(MainDocumentPart, sourceParagraphs);
                var currentById = currentMap.Blocks.ToDictionary(block => block.BlockId, StringComparer.Ordinal);
                var paragraphByPath = sourceParagraphs
                    .Select((paragraph, index) => (paragraph, block: currentMap.Blocks[index]))
                    .ToDictionary(item => item.block.StructuralPath, item => item.paragraph, StringComparer.Ordinal);

                var profile = ThesisRoleFormattingProfile.Nku2026Draft(plan.TitleMode);
                var operations = ValidateAndResolveOperations(plan, profile, currentById, paragraphByPath);
                foreach (var operation in operations)
                    ApplyTarget(operation.Paragraph, operation.Target);

                var textAfter = GetMainStoryParagraphs(body);
                var afterMap = DocumentSourceMap.Create(MainDocumentPart, textAfter);
                VerifyDocumentTextAndAnchors(currentMap, afterMap, plan.Operations);
                foreach (var operation in operations)
                    VerifyTargetApplied(operation.Paragraph, operation.Target);

                expectedParts = sourceArchive.Entries.ToDictionary(entry => entry.FullName,
                    entry => FingerprintEntry(entry), StringComparer.Ordinal);
                using var outputStream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
                using var destinationArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: false);
                foreach (var entry in sourceArchive.Entries)
                {
                    var targetEntry = destinationArchive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                    targetEntry.LastWriteTime = entry.LastWriteTime;
                    targetEntry.ExternalAttributes = entry.ExternalAttributes;
                    using var sourcePart = entry.Open();
                    using var destinationPart = targetEntry.Open();
                    if (string.Equals(entry.FullName, MainDocumentPart, StringComparison.Ordinal))
                    {
                        using var xmlWriter = XmlWriter.Create(destinationPart, new XmlWriterSettings
                        {
                            Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                            Indent = false,
                            OmitXmlDeclaration = document.Declaration is null,
                            CloseOutput = false
                        });
                        document.Save(xmlWriter);
                    }
                    else
                    {
                        sourcePart.CopyTo(destinationPart);
                    }
                }
            }

            VerifyOutputPackage(tempPath, expectedParts, plan);
            string outputHash;
            using (var outputStreamForHash = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                outputHash = Convert.ToHexString(SHA256.HashData(outputStreamForHash));
            File.Move(tempPath, fullOutputPath, overwrite: false);
            return new ThesisRoleFormatApplyResult(fullOutputPath, plan.Operations.Count, outputHash,
                plan.Omissions.Count == 0
                    ? Array.Empty<string>()
                    : new[] { $"另有 {plan.Omissions.Count} 个结构块未应用格式，保留原样；请逐项复核遗漏原因。" });
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    private static bool Matches(RoleFormatTarget target, ThesisStructureDecision decision) =>
        target.Section == decision.Section && target.Kind == decision.Kind && target.HeadingLevel == decision.HeadingLevel;

    private static void ValidateProfile(ThesisRoleFormattingProfile profile)
    {
        if (profile.Status != ThesisRoleProfileStatus.Draft)
            throw new NotSupportedException("当前原型只接受明确标记为 Draft 的配置；尚无获批规则库。");
        if (!string.Equals(profile.ProfileId, ThesisRoleFormattingProfile.Nku2026DraftProfileId, StringComparison.Ordinal) ||
            !string.Equals(profile.StandardVersion, "2026", StringComparison.Ordinal) ||
            !string.Equals(profile.Source, ThesisRoleFormattingProfile.Nku2026DraftSource, StringComparison.Ordinal))
            throw new NotSupportedException("当前原型只支持南开大学 2026 版 Draft 配置。");
        if (!Enum.IsDefined(profile.TitleMode)) throw new ArgumentOutOfRangeException(nameof(profile.TitleMode));
        if (profile.Targets is null || profile.Targets.Count == 0)
            throw new ArgumentException("格式配置没有目标项。", nameof(profile));

        var keys = new HashSet<(ThesisSectionKind, ThesisBlockKind, int?)>();
        foreach (var target in profile.Targets)
        {
            ValidateTarget(target);
            if (!keys.Add((target.Section, target.Kind, target.HeadingLevel)))
                throw new ArgumentException("格式配置存在重复的结构角色。", nameof(profile));
        }

        var canonical = ThesisRoleFormattingProfile.Nku2026Draft(profile.TitleMode);
        if (!profile.Targets.SequenceEqual(canonical.Targets))
            throw new ArgumentException("Draft profile 的目标值必须来自已审阅的显式 profile，不能由调用方补造格式值。", nameof(profile));
    }

    private static void ValidateTarget(RoleFormatTarget target)
    {
        ArgumentNullException.ThrowIfNull(target);
        if (!Enum.IsDefined(target.Section) || !Enum.IsDefined(target.Kind) ||
            (target.HeadingLevel is < 1 or > 9))
            throw new ArgumentException("格式目标含未知 section、kind 或 heading level。");
        if (target.Kind == ThesisBlockKind.Heading && target.HeadingLevel is null)
            throw new ArgumentException("标题格式目标必须指定 heading level。");
        if (target.Kind != ThesisBlockKind.Heading && target.HeadingLevel is not null)
            throw new ArgumentException("非标题格式目标不能指定 heading level。");
        if (target.EastAsiaFont is { } eastAsia && (string.IsNullOrWhiteSpace(eastAsia) || eastAsia.Length > 128) ||
            target.LatinFont is { } latin && (string.IsNullOrWhiteSpace(latin) || latin.Length > 128))
            throw new ArgumentException("字体名称为空或过长。");
        if (target.SizePt is { } size) ValidatePositiveTwips(size, "字号", max: 200, multiplier: 2);
        if (target.ExactLineSpacingPt is { } line) ValidatePositiveTwips(line, "固定行距", max: 500, multiplier: 20);
        if (target.SingleLineSpacing == false)
            throw new ArgumentException("当前 profile 只接受明确的单倍行距或未设置，不猜测其他倍数。");
        if (target.SingleLineSpacing == true && target.ExactLineSpacingPt is not null)
            throw new ArgumentException("单倍行距与固定值行距互斥。");
        if (target.BeforePt is { } before) ValidateNonNegativeTwips(before, "段前距");
        if (target.AfterPt is { } after) ValidateNonNegativeTwips(after, "段后距");
        if (target.FirstLineChars is < 0 or > 20 || target.LeftIndentChars is < 0 or > 20)
            throw new ArgumentOutOfRangeException(nameof(target), "缩进字符数必须介于 0 和 20 之间。");
        if (target.Alignment is { } alignment && !Enum.IsDefined(alignment))
            throw new ArgumentOutOfRangeException(nameof(target.Alignment));
        if (target.EastAsiaFont is null && target.LatinFont is null && target.SizePt is null && target.Bold is null &&
            target.Alignment is null && target.ExactLineSpacingPt is null && target.SingleLineSpacing is null &&
            target.BeforePt is null && target.AfterPt is null && target.FirstLineChars is null && target.LeftIndentChars is null)
            throw new ArgumentException("空格式目标没有可安全应用的显式属性。");
    }

    private static void ValidatePositiveTwips(double value, string label, double max, double multiplier)
    {
        if (!double.IsFinite(value) || value <= 0 || value > max || !IsIntegral(value * multiplier))
            throw new ArgumentOutOfRangeException(label, $"{label}必须为支持范围内且可无损编码的正值。");
    }

    private static void ValidateNonNegativeTwips(double value, string label)
    {
        if (!double.IsFinite(value) || value < 0 || value > 1000 || !IsIntegral(value * 20))
            throw new ArgumentOutOfRangeException(label, $"{label}必须为支持范围内且可无损编码的非负值。");
    }

    private static bool IsIntegral(double value) => Math.Abs(value - Math.Round(value)) < 1e-8;

    private static void ValidatePlan(ThesisRoleFormatPlan plan)
    {
        var profile = new ThesisRoleFormattingProfile(plan.ProfileId, plan.StandardVersion,
            plan.ProfileStatus, plan.TitleMode, plan.ProfileSource, ThesisRoleFormattingProfile.Nku2026Draft(plan.TitleMode).Targets);
        ValidateProfile(profile);
        if (!plan.AuthorSelectedDraft)
            throw new InvalidOperationException("计划没有作者明确选择 Draft 的记录。");
        _ = NormalizeSha256(plan.PackageSha256, nameof(plan.PackageSha256));
        if (!string.Equals(plan.SourcePartName, MainDocumentPart, StringComparison.Ordinal))
            throw new NotSupportedException("排版计划指向非主文档 part。");
        if (plan.Operations is null || plan.Omissions is null ||
            plan.Operations.Select(operation => operation.BlockId).Distinct(StringComparer.Ordinal).Count() != plan.Operations.Count)
            throw new InvalidDataException("排版计划操作列表缺失或包含重复块。");
        var knownTargets = profile.Targets;
        foreach (var operation in plan.Operations)
        {
            if (string.IsNullOrWhiteSpace(operation.BlockId) || string.IsNullOrWhiteSpace(operation.StructuralPath) ||
                string.IsNullOrWhiteSpace(operation.ExactText) || operation.ParagraphIndex < 0 ||
                !double.IsFinite(operation.Confidence) || operation.Confidence is < MinimumAutomaticConfidence or > 1 ||
                operation.RequiresReview ||
                !Enum.IsDefined(operation.Section) || !Enum.IsDefined(operation.Kind))
                throw new InvalidDataException("排版计划包含低置信度、待复核或不完整的源锚点。");
            if (!knownTargets.Any(target => Matches(target, new ThesisStructureDecision(operation.BlockId,
                    operation.ParagraphIndex, operation.Section, operation.Kind, operation.Confidence,
                    string.Empty, operation.RequiresReview, operation.HeadingLevel))))
                throw new InvalidDataException("排版计划包含不属于显式 profile 的结构角色。");
        }
    }

    private static string NormalizeSha256(string? value, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length != 64 || value.Any(character => !Uri.IsHexDigit(character)))
            throw new ArgumentException("需要完整的 SHA-256 包指纹。", parameterName);
        return value.ToUpperInvariant();
    }

    private static IReadOnlyList<ResolvedOperation> ValidateAndResolveOperations(
        ThesisRoleFormatPlan plan,
        ThesisRoleFormattingProfile profile,
        IReadOnlyDictionary<string, DocumentSourceBlock> currentById,
        IReadOnlyDictionary<string, XElement> paragraphByPath)
    {
        var resolved = new List<ResolvedOperation>(plan.Operations.Count);
        foreach (var operation in plan.Operations)
        {
            if (!currentById.TryGetValue(operation.BlockId, out var current) ||
                current.ParagraphIndex != operation.ParagraphIndex ||
                !string.Equals(current.StructuralPath, operation.StructuralPath, StringComparison.Ordinal) ||
                !string.Equals(current.ExactText, operation.ExactText, StringComparison.Ordinal) ||
                !current.HasCompleteTextMapping || current.UnsupportedReasons.Count > 0)
                throw new InvalidOperationException($"源锚点 {operation.BlockId} 的段号、路径、完整文本或映射已变化；拒绝应用。");
            if (!paragraphByPath.TryGetValue(operation.StructuralPath, out var paragraph))
                throw new InvalidOperationException($"找不到源锚点 {operation.BlockId} 的结构路径。");
            if (paragraph.Ancestors(W + "tc").Any())
                throw new InvalidOperationException($"源锚点 {operation.BlockId} 实际位于表格单元格；此原型不修改表格单元格格式。");

            var target = profile.Targets.SingleOrDefault(candidate =>
                candidate.Section == operation.Section && candidate.Kind == operation.Kind &&
                candidate.HeadingLevel == operation.HeadingLevel)
                ?? throw new InvalidDataException("排版操作没有对应的 profile 目标。");
            resolved.Add(new ResolvedOperation(operation, paragraph, target));
        }
        return resolved;
    }

    private static void ApplyTarget(XElement paragraph, RoleFormatTarget target)
    {
        if (target.Alignment is not null || target.ExactLineSpacingPt is not null || target.SingleLineSpacing is true ||
            target.BeforePt is not null || target.AfterPt is not null || target.FirstLineChars is not null ||
            target.LeftIndentChars is not null)
        {
            // w:pPr is the first child of w:p. It is not itself a pPr child property,
            // so insert it explicitly rather than passing it through the pPr ordering table.
            var pPr = paragraph.Element(W + "pPr");
            if (pPr is null)
            {
                pPr = new XElement(W + "pPr");
                paragraph.AddFirst(pPr);
            }
            if (target.Alignment is { } alignment)
            {
                var jc = GetOrCreateChild(pPr, W + "jc", ParagraphPropertyOrder);
                SetWordAttribute(jc, "val", alignment switch
                {
                    ThesisRoleAlignment.Left => "left",
                    ThesisRoleAlignment.Center => "center",
                    ThesisRoleAlignment.Right => "right",
                    ThesisRoleAlignment.Justified => "both",
                    _ => throw new ArgumentOutOfRangeException(nameof(target.Alignment))
                });
            }
            if (target.ExactLineSpacingPt is not null || target.SingleLineSpacing is true ||
                target.BeforePt is not null || target.AfterPt is not null)
            {
                var spacing = GetOrCreateChild(pPr, W + "spacing", ParagraphPropertyOrder);
                if (target.ExactLineSpacingPt is { } exactLine)
                {
                    SetWordAttribute(spacing, "line", Twips(exactLine * 20));
                    SetWordAttribute(spacing, "lineRule", "exact");
                }
                else if (target.SingleLineSpacing is true)
                {
                    SetWordAttribute(spacing, "line", "240");
                    SetWordAttribute(spacing, "lineRule", "auto");
                }
                if (target.BeforePt is { } before)
                {
                    SetWordAttribute(spacing, "before", Twips(before * 20));
                    RemoveWordAttribute(spacing, "beforeLines");
                    RemoveWordAttribute(spacing, "beforeAutospacing");
                }
                if (target.AfterPt is { } after)
                {
                    SetWordAttribute(spacing, "after", Twips(after * 20));
                    RemoveWordAttribute(spacing, "afterLines");
                    RemoveWordAttribute(spacing, "afterAutospacing");
                }
            }
            if (target.FirstLineChars is not null || target.LeftIndentChars is not null)
            {
                var indentation = GetOrCreateChild(pPr, W + "ind", ParagraphPropertyOrder);
                if (target.FirstLineChars is { } firstLine)
                {
                    SetWordAttribute(indentation, "firstLineChars", (firstLine * 100).ToString(System.Globalization.CultureInfo.InvariantCulture));
                    RemoveWordAttribute(indentation, "firstLine");
                    RemoveWordAttribute(indentation, "hanging");
                    RemoveWordAttribute(indentation, "hangingChars");
                }
                if (target.LeftIndentChars is { } left)
                {
                    SetWordAttribute(indentation, "leftChars", (left * 100).ToString(System.Globalization.CultureInfo.InvariantCulture));
                    RemoveWordAttribute(indentation, "left");
                }
            }
        }

        foreach (var run in paragraph.Elements(W + "r").Where(HasText))
        {
            if (target.EastAsiaFont is null && target.LatinFont is null && target.SizePt is null && target.Bold is null)
                break;
            // w:rPr must precede every content child inside w:r.
            var rPr = run.Element(W + "rPr");
            if (rPr is null)
            {
                rPr = new XElement(W + "rPr");
                run.AddFirst(rPr);
            }
            if (target.EastAsiaFont is not null || target.LatinFont is not null)
            {
                var fonts = GetOrCreateChild(rPr, W + "rFonts", RunPropertyOrder);
                if (target.EastAsiaFont is { } eastAsia)
                {
                    SetWordAttribute(fonts, "eastAsia", eastAsia);
                    RemoveWordAttribute(fonts, "eastAsiaTheme");
                }
                if (target.LatinFont is { } latin)
                {
                    SetWordAttribute(fonts, "ascii", latin);
                    SetWordAttribute(fonts, "hAnsi", latin);
                    RemoveWordAttribute(fonts, "asciiTheme");
                    RemoveWordAttribute(fonts, "hAnsiTheme");
                }
            }
            if (target.Bold is { } bold)
            {
                SetBooleanProperty(GetOrCreateChild(rPr, W + "b", RunPropertyOrder), bold);
                SetBooleanProperty(GetOrCreateChild(rPr, W + "bCs", RunPropertyOrder), bold);
            }
            if (target.SizePt is { } size)
            {
                var halfPoints = Twips(size * 2);
                SetWordAttribute(GetOrCreateChild(rPr, W + "sz", RunPropertyOrder), "val", halfPoints);
                SetWordAttribute(GetOrCreateChild(rPr, W + "szCs", RunPropertyOrder), "val", halfPoints);
            }
        }
    }

    private static bool HasText(XElement run) => run.Descendants(W + "t").Any();

    private static void SetBooleanProperty(XElement element, bool value)
    {
        if (value) RemoveWordAttribute(element, "val");
        else SetWordAttribute(element, "val", "0");
    }

    private static XElement GetOrCreateChild(XElement parent, XName name, IReadOnlyDictionary<string, int> order)
    {
        var matches = parent.Elements(name).ToArray();
        if (matches.Length > 1) throw new InvalidDataException($"OOXML contains duplicate {name.LocalName} properties.");
        if (matches.Length == 1) return matches[0];
        var child = new XElement(name);
        if (!order.TryGetValue(name.LocalName, out var targetOrder))
        {
            parent.Add(child);
            return child;
        }
        var later = parent.Elements().FirstOrDefault(sibling =>
            order.TryGetValue(sibling.Name.LocalName, out var siblingOrder) && siblingOrder > targetOrder);
        if (later is null) parent.Add(child);
        else later.AddBeforeSelf(child);
        return child;
    }

    private static void SetWordAttribute(XElement element, string localName, string value) =>
        element.SetAttributeValue(W + localName, value);

    private static void RemoveWordAttribute(XElement element, string localName) =>
        element.Attribute(W + localName)?.Remove();

    private static string Twips(double value) =>
        checked((int)Math.Round(value, MidpointRounding.AwayFromZero)).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static XElement[] GetMainStoryParagraphs(XElement body) => body.Descendants(W + "p")
        .Where(paragraph => !paragraph.Ancestors(W + "txbxContent").Any())
        .ToArray();

    private static void VerifyDocumentTextAndAnchors(
        DocumentSourceMap before,
        DocumentSourceMap after,
        IReadOnlyList<ThesisRoleFormatOperation> operations)
    {
        if (before.Blocks.Count != after.Blocks.Count)
            throw new InvalidDataException("排版后段落数量变化。");
        for (var index = 0; index < before.Blocks.Count; index++)
        {
            var left = before.Blocks[index];
            var right = after.Blocks[index];
            if (!string.Equals(left.StructuralPath, right.StructuralPath, StringComparison.Ordinal) ||
                !string.Equals(left.ExactText, right.ExactText, StringComparison.Ordinal))
                throw new InvalidDataException($"排版改变了原文或结构锚点（段落 {index}）。");
        }
        foreach (var operation in operations)
        {
            if (!after.TryGetBlock(operation.BlockId, out var anchor) || anchor is null ||
                anchor.ParagraphIndex != operation.ParagraphIndex ||
                !string.Equals(anchor.StructuralPath, operation.StructuralPath, StringComparison.Ordinal) ||
                !string.Equals(anchor.ExactText, operation.ExactText, StringComparison.Ordinal))
                throw new InvalidDataException($"排版后源锚点 {operation.BlockId} 无法再次验证。");
        }
    }

    private static void VerifyTargetApplied(XElement paragraph, RoleFormatTarget target)
    {
        var pPr = paragraph.Element(W + "pPr");
        if (target.Alignment is { } alignment)
        {
            var expected = alignment switch
            {
                ThesisRoleAlignment.Left => "left",
                ThesisRoleAlignment.Center => "center",
                ThesisRoleAlignment.Right => "right",
                ThesisRoleAlignment.Justified => "both",
                _ => throw new ArgumentOutOfRangeException(nameof(target.Alignment))
            };
            if ((string?)pPr?.Element(W + "jc")?.Attribute(W + "val") != expected)
                throw new InvalidDataException("排版后对齐方式没有读回目标值。");
        }
        var spacing = pPr?.Element(W + "spacing");
        if (target.ExactLineSpacingPt is { } exactLine &&
            ((string?)spacing?.Attribute(W + "line") != Twips(exactLine * 20) ||
             (string?)spacing?.Attribute(W + "lineRule") != "exact"))
            throw new InvalidDataException("排版后固定行距没有读回目标值。");
        if (target.SingleLineSpacing is true &&
            ((string?)spacing?.Attribute(W + "line") != "240" || (string?)spacing?.Attribute(W + "lineRule") != "auto"))
            throw new InvalidDataException("排版后单倍行距没有读回目标值。");
        if (target.BeforePt is { } before && (string?)spacing?.Attribute(W + "before") != Twips(before * 20))
            throw new InvalidDataException("排版后段前距没有读回目标值。");
        if (target.AfterPt is { } after && (string?)spacing?.Attribute(W + "after") != Twips(after * 20))
            throw new InvalidDataException("排版后段后距没有读回目标值。");
        var indentation = pPr?.Element(W + "ind");
        if (target.FirstLineChars is { } firstLine &&
            (string?)indentation?.Attribute(W + "firstLineChars") != (firstLine * 100).ToString(System.Globalization.CultureInfo.InvariantCulture))
            throw new InvalidDataException("排版后首行字符缩进没有读回目标值。");
        if (target.LeftIndentChars is { } left &&
            (string?)indentation?.Attribute(W + "leftChars") != (left * 100).ToString(System.Globalization.CultureInfo.InvariantCulture))
            throw new InvalidDataException("排版后左侧字符缩进没有读回目标值。");

        foreach (var run in paragraph.Elements(W + "r").Where(HasText))
        {
            var rPr = run.Element(W + "rPr");
            var fonts = rPr?.Element(W + "rFonts");
            if (target.EastAsiaFont is { } eastAsia && (string?)fonts?.Attribute(W + "eastAsia") != eastAsia)
                throw new InvalidDataException("排版后东亚字体没有读回目标值。");
            if (target.LatinFont is { } latin &&
                ((string?)fonts?.Attribute(W + "ascii") != latin || (string?)fonts?.Attribute(W + "hAnsi") != latin))
                throw new InvalidDataException("排版后拉丁字体没有读回目标值。");
            if (target.SizePt is { } size)
            {
                var expectedSize = Twips(size * 2);
                if ((string?)rPr?.Element(W + "sz")?.Attribute(W + "val") != expectedSize ||
                    (string?)rPr?.Element(W + "szCs")?.Attribute(W + "val") != expectedSize)
                    throw new InvalidDataException("排版后字号没有读回目标值。");
            }
            if (target.Bold is { } bold)
            {
                var expected = bold ? null : "0";
                if ((string?)rPr?.Element(W + "b")?.Attribute(W + "val") != expected ||
                    (string?)rPr?.Element(W + "bCs")?.Attribute(W + "val") != expected)
                    throw new InvalidDataException("排版后加粗属性没有读回目标值。");
            }
        }
    }

    private static XDocument LoadXml(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaximumDocumentXmlBytes,
            IgnoreWhitespace = false
        };
        using var reader = XmlReader.Create(stream, settings);
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static void ValidateArchiveEntries(ZipArchive archive)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        long uncompressedTotal = 0;
        if (archive.Entries.Count > 20_000)
            throw new InvalidDataException("DOCX 包含过多 part，超出原型支持范围。");
        foreach (var entry in archive.Entries)
        {
            if (!names.Add(entry.FullName)) throw new InvalidDataException("DOCX 包含重复 part 名称。");
            if (entry.FullName.StartsWith('/') || entry.FullName.Contains('\\') ||
                entry.FullName.Split('/').Any(segment => segment is "." or ".."))
                throw new InvalidDataException("DOCX 包含不安全的 part 路径。");
            if (entry.Length < 0 || entry.Length > MaximumPackageBytes)
                throw new InvalidDataException("DOCX 包含超出支持范围的 part。");
            uncompressedTotal = checked(uncompressedTotal + entry.Length);
            if (uncompressedTotal > MaximumPackageBytes)
                throw new InvalidDataException("DOCX 解压后总大小超过原型支持范围。");
        }
    }

    private static ZipArchiveEntry FindUniqueEntry(ZipArchive archive, string name)
    {
        var matches = archive.Entries.Where(entry => string.Equals(entry.FullName, name, StringComparison.Ordinal)).ToArray();
        return matches.Length == 1 ? matches[0] : throw new InvalidDataException($"DOCX 必须恰有一个 {name} part。");
    }

    private static PartFingerprint FingerprintEntry(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return new PartFingerprint(entry.Length, Convert.ToHexString(SHA256.HashData(stream)));
    }

    private static void VerifyOutputPackage(
        string path,
        IReadOnlyDictionary<string, PartFingerprint> expectedParts,
        ThesisRoleFormatPlan plan)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        ValidateArchiveEntries(archive);
        if (archive.Entries.Count != expectedParts.Count)
            throw new InvalidDataException("另存文件中的 OOXML part 数量变化。");
        foreach (var entry in archive.Entries)
        {
            if (!expectedParts.TryGetValue(entry.FullName, out var expected))
                throw new InvalidDataException($"另存文件出现未预期的 OOXML part {entry.FullName}。");
            if (string.Equals(entry.FullName, MainDocumentPart, StringComparison.Ordinal)) continue;
            var actual = FingerprintEntry(entry);
            if (actual.Length != expected.Length || !string.Equals(actual.Sha256, expected.Sha256, StringComparison.Ordinal))
                throw new InvalidDataException($"非文档 part {entry.FullName} 内容发生变化。");
        }

        var document = LoadXml(FindUniqueEntry(archive, MainDocumentPart));
        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidDataException("另存文件缺少主文档 body。");
        var paragraphs = GetMainStoryParagraphs(body);
        var map = DocumentSourceMap.Create(MainDocumentPart, paragraphs);
        var paragraphByPath = paragraphs.Select((paragraph, index) => (paragraph, block: map.Blocks[index]))
            .ToDictionary(item => item.block.StructuralPath, item => item.paragraph, StringComparer.Ordinal);
        foreach (var operation in plan.Operations)
        {
            if (!map.TryGetBlock(operation.BlockId, out var block) || block is null ||
                block.ParagraphIndex != operation.ParagraphIndex ||
                !string.Equals(block.StructuralPath, operation.StructuralPath, StringComparison.Ordinal) ||
                !string.Equals(block.ExactText, operation.ExactText, StringComparison.Ordinal) ||
                !paragraphByPath.TryGetValue(operation.StructuralPath, out var paragraph))
                throw new InvalidDataException($"另存文件的源锚点 {operation.BlockId} 校验失败。");
            var target = ThesisRoleFormattingProfile.Nku2026Draft(plan.TitleMode).Targets.Single(candidate =>
                candidate.Section == operation.Section && candidate.Kind == operation.Kind &&
                candidate.HeadingLevel == operation.HeadingLevel);
            VerifyTargetApplied(paragraph, target);
        }
    }

    private sealed record PartFingerprint(long Length, string Sha256);
    private sealed record ResolvedOperation(ThesisRoleFormatOperation Operation, XElement Paragraph, RoleFormatTarget Target);
}
