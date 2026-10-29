using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ThesisAddIn.Protocol;

// Kept independent of ThesisEngine.Core so the same DTOs load in .NET Framework Word.
public enum EngineThesisSection { Unknown, FrontMatter, ChineseAbstract, EnglishAbstract, Contents, Body, References, Appendix, Acknowledgments }
public enum EngineThesisBlockKind { Unknown, Paragraph, Heading, Keywords, ContentsEntry, FigureCaption, TableCaption, TableCell }
public enum EngineThesisTitleMode { Mode1, Mode2 }
public enum EngineThesisRoleAlignment { Left, Center, Right, Justified }
public enum EngineThesisProfileStatus { Draft, Approved, Custom }

public class EngineStructureSourceBlock
{
    [JsonPropertyName("blockId")] public string BlockId { get; set; } = string.Empty;
    [JsonPropertyName("partName")] public string PartName { get; set; } = "word/document.xml";
    [JsonPropertyName("paragraphIndex")] public int ParagraphIndex { get; set; }
    /// <summary>One-based Word Paragraphs index, including completed-table-row CP compensation.</summary>
    [JsonPropertyName("wordParagraphIndex")] public int WordParagraphIndex { get; set; }
    [JsonPropertyName("location")] public string Location { get; set; } = string.Empty;
    [JsonPropertyName("firstRunLocation")] public string? FirstRunLocation { get; set; }
    /// <summary>Only nonempty direct text runs; no paragraph mark or hidden comment reference run.</summary>
    [JsonPropertyName("runLocations")] public List<string> RunLocations { get; set; } = new();
    [JsonPropertyName("structuralPath")] public string StructuralPath { get; set; } = string.Empty;
    [JsonPropertyName("exactText")] public string ExactText { get; set; } = string.Empty;
    [JsonPropertyName("textHash")] public string TextHash { get; set; } = string.Empty;
    [JsonPropertyName("isTableCell")] public bool IsTableCell { get; set; }
    [JsonPropertyName("hasCompleteTextMapping")] public bool HasCompleteTextMapping { get; set; }
}

public sealed class EngineStructureBlock : EngineStructureSourceBlock
{
    [JsonPropertyName("section")] public EngineThesisSection Section { get; set; }
    [JsonPropertyName("kind")] public EngineThesisBlockKind Kind { get; set; }
    [JsonPropertyName("headingLevel")] public int? HeadingLevel { get; set; }
    [JsonPropertyName("confidence")] public double Confidence { get; set; }
    [JsonPropertyName("requiresReview")] public bool RequiresReview { get; set; }
    [JsonPropertyName("evidence")] public string Evidence { get; set; } = string.Empty;
    /// <summary>Physical source can safely be confirmed, even when its inferred role is unknown.</summary>
    [JsonPropertyName("canConfirm")] public bool CanConfirm { get; set; }
}

public sealed class EngineStructureReviewResult
{
    [JsonPropertyName("documentFingerprint")] public string DocumentFingerprint { get; set; } = string.Empty;
    [JsonPropertyName("packageSha256")] public string PackageSha256 { get; set; } = string.Empty;
    [JsonPropertyName("sourcePartName")] public string SourcePartName { get; set; } = "word/document.xml";
    [JsonPropertyName("sourceParagraphCount")] public int SourceParagraphCount { get; set; }
    [JsonPropertyName("totalTableRows")] public int TotalTableRows { get; set; }
    [JsonPropertyName("expectedWordParagraphCount")] public int ExpectedWordParagraphCount { get; set; }
    [JsonPropertyName("blocks")] public List<EngineStructureBlock> Blocks { get; set; } = new();
    [JsonPropertyName("warnings")] public List<string> Warnings { get; set; } = new();
}

public sealed class EngineApprovedStructureBlock
{
    [JsonPropertyName("blockId")] public string BlockId { get; set; } = string.Empty;
    [JsonPropertyName("section")] public EngineThesisSection Section { get; set; }
    [JsonPropertyName("kind")] public EngineThesisBlockKind Kind { get; set; }
    [JsonPropertyName("headingLevel")] public int? HeadingLevel { get; set; }
}

public sealed class EngineStructurePlanRequest
{
    [JsonPropertyName("documentFingerprint")] public string DocumentFingerprint { get; set; } = string.Empty;
    [JsonPropertyName("packageSha256")] public string PackageSha256 { get; set; } = string.Empty;
    [JsonPropertyName("approvedBlocks")] public List<EngineApprovedStructureBlock> ApprovedBlocks { get; set; } = new();
    [JsonPropertyName("titleMode")] public EngineThesisTitleMode TitleMode { get; set; }
    [JsonPropertyName("authorSelectedDraft")] public bool AuthorSelectedDraft { get; set; }
}

/// <summary>Canonical profile values, not model-generated properties or executable commands.</summary>
public sealed class EngineStructureFormatTarget
{
    [JsonPropertyName("eastAsiaFont")] public string? EastAsiaFont { get; set; }
    [JsonPropertyName("latinFont")] public string? LatinFont { get; set; }
    [JsonPropertyName("sizePt")] public double? SizePt { get; set; }
    [JsonPropertyName("bold")] public bool? Bold { get; set; }
    [JsonPropertyName("alignment")] public EngineThesisRoleAlignment? Alignment { get; set; }
    [JsonPropertyName("exactLineSpacingPt")] public double? ExactLineSpacingPt { get; set; }
    [JsonPropertyName("singleLineSpacing")] public bool? SingleLineSpacing { get; set; }
    [JsonPropertyName("beforePt")] public double? BeforePt { get; set; }
    [JsonPropertyName("afterPt")] public double? AfterPt { get; set; }
    [JsonPropertyName("firstLineChars")] public int? FirstLineChars { get; set; }
    [JsonPropertyName("leftIndentChars")] public int? LeftIndentChars { get; set; }
}

public sealed class EngineStructureFormatOperation : EngineStructureSourceBlock
{
    [JsonPropertyName("section")] public EngineThesisSection Section { get; set; }
    [JsonPropertyName("kind")] public EngineThesisBlockKind Kind { get; set; }
    [JsonPropertyName("headingLevel")] public int? HeadingLevel { get; set; }
    [JsonPropertyName("confidence")] public double Confidence { get; set; }
    [JsonPropertyName("requiresReview")] public bool RequiresReview { get; set; }
    [JsonPropertyName("target")] public EngineStructureFormatTarget Target { get; set; } = new();
}

public sealed class EngineStructureOmission
{
    [JsonPropertyName("blockId")] public string? BlockId { get; set; }
    [JsonPropertyName("reason")] public string Reason { get; set; } = string.Empty;
}

public sealed class EngineStructureFormatPlan
{
    [JsonPropertyName("documentFingerprint")] public string DocumentFingerprint { get; set; } = string.Empty;
    [JsonPropertyName("packageSha256")] public string PackageSha256 { get; set; } = string.Empty;
    [JsonPropertyName("sourcePartName")] public string SourcePartName { get; set; } = "word/document.xml";
    [JsonPropertyName("sourceParagraphCount")] public int SourceParagraphCount { get; set; }
    [JsonPropertyName("totalTableRows")] public int TotalTableRows { get; set; }
    [JsonPropertyName("expectedWordParagraphCount")] public int ExpectedWordParagraphCount { get; set; }
    [JsonPropertyName("profileId")] public string ProfileId { get; set; } = string.Empty;
    [JsonPropertyName("standardVersion")] public string StandardVersion { get; set; } = string.Empty;
    [JsonPropertyName("profileStatus")] public EngineThesisProfileStatus ProfileStatus { get; set; }
    [JsonPropertyName("titleMode")] public EngineThesisTitleMode TitleMode { get; set; }
    [JsonPropertyName("profileSource")] public string ProfileSource { get; set; } = string.Empty;
    [JsonPropertyName("authorSelectedDraft")] public bool AuthorSelectedDraft { get; set; }
    [JsonPropertyName("sourceBlocks")] public List<EngineStructureSourceBlock> SourceBlocks { get; set; } = new();
    [JsonPropertyName("operations")] public List<EngineStructureFormatOperation> Operations { get; set; } = new();
    [JsonPropertyName("omissions")] public List<EngineStructureOmission> Omissions { get; set; } = new();
    [JsonPropertyName("warnings")] public List<string> Warnings { get; set; } = new();
}

public sealed class EngineStructureVerificationRequest
{
    [JsonPropertyName("previousPlan")] public EngineStructureFormatPlan PreviousPlan { get; set; } = new();
}

public sealed class EngineStructureVerificationIssue
{
    [JsonPropertyName("blockId")] public string? BlockId { get; set; }
    [JsonPropertyName("paragraphIndex")] public int? ParagraphIndex { get; set; }
    [JsonPropertyName("property")] public string Property { get; set; } = string.Empty;
    [JsonPropertyName("expected")] public string Expected { get; set; } = string.Empty;
    [JsonPropertyName("actual")] public string Actual { get; set; } = string.Empty;
    [JsonPropertyName("reason")] public string Reason { get; set; } = string.Empty;
}

public sealed class EngineStructureVerificationResult
{
    [JsonPropertyName("succeeded")] public bool Succeeded { get; set; }
    [JsonPropertyName("textUnchanged")] public bool TextUnchanged { get; set; }
    [JsonPropertyName("targetsMatch")] public bool TargetsMatch { get; set; }
    [JsonPropertyName("verifiedOperations")] public int VerifiedOperations { get; set; }
    [JsonPropertyName("issues")] public List<EngineStructureVerificationIssue> Issues { get; set; } = new();
    [JsonPropertyName("warnings")] public List<string> Warnings { get; set; } = new();
    [JsonPropertyName("omissions")] public List<EngineStructureOmission> Omissions { get; set; } = new();
}
