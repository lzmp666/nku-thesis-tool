using ThesisEngine.Core.Rules;

namespace ThesisEngine.Core.Analysis;

public sealed record MappedParagraph(
    int ParagraphIndex,
    string ObjectKind,
    string ObjectLabel,
    ParagraphSnapshot Paragraph,
    ThesisStructureDecision? Decision,
    ParagraphLayout? Layout,
    bool Checkable,
    string? UncheckReason)
{
    public double Confidence => Decision?.Confidence ?? 0;
}

/// <summary>
/// Reuses the structure classifier's block decisions. It deliberately does not
/// re-implement section/heading detection; it only maps those decisions to the
/// object kinds used by the YAML rules.
/// </summary>
public static class ThesisObjectMap
{
    public const string TableCell = "table_cell";
    public const string Unknown = "unknown";

    public static IReadOnlyList<MappedParagraph> Create(DocumentSnapshot snapshot, DocumentLayoutProbe layout)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(layout);
        var decisions = snapshot.Structure?.Blocks
            .GroupBy(block => block.ParagraphIndex)
            .ToDictionary(group => group.Key, group => group.First())
            ?? new Dictionary<int, ThesisStructureDecision>();

        var result = new List<MappedParagraph>();
        foreach (var paragraph in snapshot.Paragraphs)
        {
            if (string.IsNullOrWhiteSpace(paragraph.Text) && paragraph.Caption is null && !paragraph.IsTableCell)
                continue;

            decisions.TryGetValue(paragraph.ParagraphIndex, out var decision);
            var (kind, label) = Classify(paragraph, decision);
            var paragraphLayout = layout.TryGetParagraph(paragraph.ParagraphIndex);
            var (checkable, reason) = Checkability(paragraph, decision, paragraphLayout, kind);
            result.Add(new MappedParagraph(
                paragraph.ParagraphIndex,
                kind,
                label,
                paragraph,
                decision,
                paragraphLayout,
                checkable,
                reason));
        }
        return result.AsReadOnly();
    }

    private static (string Kind, string Label) Classify(ParagraphSnapshot paragraph, ThesisStructureDecision? decision)
    {
        if (paragraph.IsTableCell) return (TableCell, "表格单元格");
        if (paragraph.Caption == CaptionKind.Figure) return (AnalysisObjectKinds.Caption, "图题");
        if (paragraph.Caption == CaptionKind.Table) return (AnalysisObjectKinds.Caption, "表题");
        if (decision is null) return (Unknown, "未分类段落");

        return (decision.Section, decision.Kind) switch
        {
            (ThesisSectionKind.Body, ThesisBlockKind.Heading) => HeadingKind(decision.HeadingLevel),
            (ThesisSectionKind.Body, ThesisBlockKind.Paragraph) => (AnalysisObjectKinds.BodyText, "正文段落"),
            (ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Heading) => (AnalysisObjectKinds.AbstractCnTitle, "中文摘要标题"),
            (ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Paragraph or ThesisBlockKind.Keywords) =>
                (AnalysisObjectKinds.AbstractCnBody, "中文摘要正文"),
            (ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Heading) => (AnalysisObjectKinds.AbstractEnTitle, "英文摘要标题"),
            (ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Paragraph or ThesisBlockKind.Keywords) =>
                (AnalysisObjectKinds.AbstractEnBody, "英文摘要正文"),
            (ThesisSectionKind.Contents, ThesisBlockKind.Heading) => (AnalysisObjectKinds.TocTitle, "目录标题"),
            (ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry) => (AnalysisObjectKinds.TocEntry, "目录条目"),
            (ThesisSectionKind.References, ThesisBlockKind.Heading) => (AnalysisObjectKinds.HeadingChapter, "参考文献标题"),
            (ThesisSectionKind.References, ThesisBlockKind.Paragraph) => (AnalysisObjectKinds.ReferencesBody, "参考文献正文"),
            (ThesisSectionKind.Appendix, ThesisBlockKind.Heading) => (AnalysisObjectKinds.HeadingChapter, "附录标题"),
            (ThesisSectionKind.Appendix, ThesisBlockKind.Paragraph) => (AnalysisObjectKinds.AppendixBody, "附录正文"),
            (ThesisSectionKind.Acknowledgments, ThesisBlockKind.Heading) => (AnalysisObjectKinds.HeadingChapter, "致谢标题"),
            (ThesisSectionKind.Acknowledgments, ThesisBlockKind.Paragraph) =>
                (AnalysisObjectKinds.AcknowledgementBody, "致谢正文"),
            (ThesisSectionKind.FrontMatter, _) => (AnalysisObjectKinds.Cover, "封面/题名页"),
            _ => (Unknown, "未分类段落")
        };
    }

    private static (string Kind, string Label) HeadingKind(int? level) => level switch
    {
        1 => (AnalysisObjectKinds.HeadingChapter, "章标题"),
        2 => (AnalysisObjectKinds.HeadingLevel1, "一级标题"),
        3 => (AnalysisObjectKinds.HeadingLevel2, "二级标题"),
        4 => (AnalysisObjectKinds.HeadingLevel3, "三级标题"),
        _ => (Unknown, "未识别层级的标题")
    };

    private static (bool Checkable, string? Reason) Checkability(
        ParagraphSnapshot paragraph,
        ThesisStructureDecision? decision,
        ParagraphLayout? layout,
        string kind)
    {
        if (kind is TableCell or AnalysisObjectKinds.Caption)
            return (false, "图表/表格规则由 #92 的 FigureTable* 处理，本分析器不重复检查");
        if (kind == AnalysisObjectKinds.Cover)
            return (false, "封面/题名页不在 v1 检查范围（RFC-6）");
        if (kind == AnalysisObjectKinds.TocEntry)
            return (false, "目录条目检测对象（TOC 域样式 vs 渲染结果）待定（T6）");
        if (kind == Unknown)
            return (false, "结构分类无法确认");
        if (decision is null)
            return (false, "缺少结构分类结果");
        if (paragraph.HasUnsupportedInlineContent)
            return (false, "段落含字段、修订或其他未支持的内联结构");
        if (paragraph.ParagraphFormatting?.HasUnresolvedStyle == true)
            return (false, "段落样式链无法解析");
        if (layout is null)
            return (false, "段落格式探测缺失");
        if (!layout.ValuesKnown)
            return (false, "段落样式链无法解析");
        if (decision.RequiresReview)
            return (false, $"结构分类需要复核（置信度 {decision.Confidence:0.00}）");
        if (decision.Confidence < 0.75)
            return (false, $"结构分类置信度过低（{decision.Confidence:0.00}）");
        return (true, null);
    }
}
