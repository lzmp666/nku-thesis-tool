namespace ThesisEngine.Core;

public sealed record ApprovedThesisBlock(string BlockId, ThesisSectionKind Section, ThesisBlockKind Kind,
    int? HeadingLevel = null);
public sealed record ThesisOutlineApproval(string DocumentFingerprint, IReadOnlyList<ApprovedThesisBlock> Blocks);
public sealed record ThesisStructureTrainingExample(string DocumentGroup, string BlockId, string Text,
    ThesisSectionKind Section, ThesisBlockKind Kind, int? HeadingLevel, string LabelSource);

/// <summary>Only an explicit author outline confirmation can promote model proposals to usable roles.</summary>
public static class ThesisOutlineReview
{
    public static DocumentSnapshot Apply(DocumentSnapshot snapshot, ThesisOutlineApproval approval)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(approval);
        if (snapshot.SourceMap is null || snapshot.Structure is null || approval.Blocks is null ||
            approval.DocumentFingerprint != ThesisStructureAiCoordinator.Fingerprint(snapshot))
            throw new InvalidDataException("大纲确认不属于当前文档快照，请重新生成结构报告。");
        var byId = snapshot.Structure.Blocks.ToDictionary(block => block.BlockId, StringComparer.Ordinal);
        var sources = snapshot.SourceMap.Blocks.ToDictionary(block => block.BlockId, StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var approved in approval.Blocks)
        {
            if (approved is null || string.IsNullOrWhiteSpace(approved.BlockId) || !seen.Add(approved.BlockId) ||
                !sources.TryGetValue(approved.BlockId, out var source) || !byId.TryGetValue(approved.BlockId, out var original))
                throw new InvalidDataException("确认大纲包含未知或重复的文本块。");
            var paragraph = snapshot.Paragraphs[source.ParagraphIndex];
            if (!source.HasCompleteTextMapping || paragraph.HasUnsupportedInlineContent || paragraph.IsTableCell ||
                !Enum.IsDefined(approved.Section) || !Enum.IsDefined(approved.Kind) ||
                approved.Kind == ThesisBlockKind.TableCell ||
                approved.Kind == ThesisBlockKind.Heading && approved.HeadingLevel is not (>= 1 and <= 9) ||
                approved.Kind != ThesisBlockKind.Heading && approved.HeadingLevel is not null)
                throw new InvalidDataException("确认大纲不能改变物理表格结构、越过不支持内容或使用无效类别/标题层级。");
            if (paragraph.Caption is not null && approved.Kind != original.Kind)
                throw new InvalidDataException("确认大纲不能把已配对题注改成正文或其他块类型。");
            var unknown = approved.Section == ThesisSectionKind.Unknown || approved.Kind == ThesisBlockKind.Unknown;
            byId[approved.BlockId] = original with
            {
                Section = approved.Section, Kind = approved.Kind, HeadingLevel = approved.HeadingLevel,
                Confidence = unknown ? 0 : 1, RequiresReview = unknown,
                Evidence = unknown ? "作者保留未知结构。" : "作者已明确确认此文本块的大纲角色；分数表示确认状态。"
            };
        }
        var structure = snapshot.Structure with { Blocks = snapshot.Structure.Blocks.Select(block => byId[block.BlockId]).ToArray() };
        var paragraphs = snapshot.Paragraphs.Select(paragraph =>
        {
            if (paragraph.BlockId is null || !seen.Contains(paragraph.BlockId)) return paragraph;
            var decision = byId[paragraph.BlockId];
            var scope = decision.RequiresReview ? ParagraphScopeClassification.Unknown
                : decision.Section == ThesisSectionKind.Body && decision.Kind == ThesisBlockKind.Paragraph
                    ? ParagraphScopeClassification.ConfirmedBodyProse : ParagraphScopeClassification.Excluded;
            return paragraph with { Structure = decision, Scope = scope, ScopeReason = decision.Evidence };
        }).ToArray();
        return snapshot with { Paragraphs = paragraphs, Structure = structure };
    }

    public static IReadOnlyList<ThesisStructureTrainingExample> ExportApprovedSamples(
        DocumentSnapshot snapshot, ThesisOutlineApproval approval)
    {
        var confirmed = Apply(snapshot, approval); // Apply all input validation before exporting a label.
        var ids = approval.Blocks.Select(block => block.BlockId).ToHashSet(StringComparer.Ordinal);
        var source = confirmed.SourceMap!.Blocks.ToDictionary(block => block.BlockId, StringComparer.Ordinal);
        return confirmed.Structure!.Blocks.Where(block => ids.Contains(block.BlockId) && !block.RequiresReview)
            .Select(block => new ThesisStructureTrainingExample(approval.DocumentFingerprint, block.BlockId,
                source[block.BlockId].ExactText, block.Section, block.Kind, block.HeadingLevel, "author-confirmed"))
            .ToArray();
    }
}
