namespace ThesisEngine.Core.Rules;

public static class AnalysisObjectKinds
{
    public const string Page = "page";
    public const string Header = "header";
    public const string PageNumber = "page_number";
    public const string AbstractCnTitle = "abstract_cn.title";
    public const string AbstractCnBody = "abstract_cn.body";
    public const string AbstractEnTitle = "abstract_en.title";
    public const string AbstractEnBody = "abstract_en.body";
    public const string TocTitle = "toc.title";
    public const string TocEntry = "toc.entry";
    public const string HeadingChapter = "heading.chapter";
    public const string HeadingLevel1 = "heading.level1";
    public const string HeadingLevel2 = "heading.level2";
    public const string HeadingLevel3 = "heading.level3";
    public const string BodyText = "body.text";
    public const string ReferencesBody = "references.body";
    public const string AppendixBody = "appendix.body";
    public const string AcknowledgementBody = "acknowledgement.body";
    public const string SymbolTable = "symbol_table";
    public const string Caption = "caption";
    public const string Equation = "equation";
    public const string Cover = "cover";
}

public static class RuleProperties
{
    public const string FontEastAsia = "font.east_asian";
    public const string FontLatin = "font.latin";
    public const string FontSize = "font_size";
    public const string Alignment = "alignment";
    public const string LineSpacing = "line_spacing";
    public const string SpaceBefore = "space_before";
    public const string SpaceAfter = "space_after";
    public const string FirstLineIndent = "first_line_indent";
    public const string Bold = "bold";
    public const string KeywordBold = "keyword_bold";
    public const string PagePaper = "page_paper";
    public const string PageMargins = "page_margins";
    public const string PageHeaderFooter = "page_header_footer";
    public const string PageGutter = "page_gutter";
    public const string Numbering = "numbering";
    public const string TocEntries = "toc_entries";
    public const string HeadingReuse = "heading_reuse";
    public const string Caption = "caption";
    public const string Equation = "equation";
    public const string Cover = "cover";
    public const string Unknown = "unknown";
}

internal sealed record LegacySlot(
    string GroupPath,
    string RelativePath,
    string ObjectKind,
    string Property,
    bool Implemented = true,
    string? UncheckedReason = null,
    string? ExpectedOverride = null,
    string Severity = "warning",
    string FixLevel = "manual");

/// <summary>
/// Maps the reviewed legacy `general.yaml` groups to analyser targets. The mapping
/// is explicit so a newly added YAML field cannot silently turn into a pass.
/// </summary>
internal static class LegacyRuleCatalog
{
    public static IReadOnlyList<LegacySlot> Slots { get; } = Build();

    private static IReadOnlyList<LegacySlot> Build()
    {
        var slots = new List<LegacySlot>
        {
            // R1 page setup.
            new("page_setup", "paper_size", AnalysisObjectKinds.Page, RuleProperties.PagePaper),
            new("page_setup", "margins_cm", AnalysisObjectKinds.Page, RuleProperties.PageMargins),
            new("page_setup", "header_footer_cm", AnalysisObjectKinds.Page, RuleProperties.PageHeaderFooter,
                false, "页眉/页脚距边界的 UI 对应关系待确认（T5）"),
            new("page_setup", "gutter_cm", AnalysisObjectKinds.Page, RuleProperties.PageGutter,
                false, "装订线检测未实现（T5）"),

            // R2 header and R3 page number are not present in the main story reader.
            new("header", "font_cn", AnalysisObjectKinds.Header, RuleProperties.FontEastAsia,
                false, "分析器只读取主文档，页眉内容未检查"),
            new("header", "font_western", AnalysisObjectKinds.Header, RuleProperties.FontLatin,
                false, "分析器只读取主文档，页眉内容未检查"),
            new("header", "size_pt", AnalysisObjectKinds.Header, RuleProperties.FontSize,
                false, "分析器只读取主文档，页眉内容未检查"),
            new("header", "alignment", AnalysisObjectKinds.Header, RuleProperties.Alignment,
                false, "分析器只读取主文档，页眉内容未检查"),
            new("page_number", "font_cn", AnalysisObjectKinds.PageNumber, RuleProperties.FontEastAsia,
                false, "页码字段与页脚未读取"),
            new("page_number", "size_pt", AnalysisObjectKinds.PageNumber, RuleProperties.FontSize,
                false, "页码字段与页脚未读取"),
            new("page_number", "alignment", AnalysisObjectKinds.PageNumber, RuleProperties.Alignment,
                false, "页码字段与页脚未读取"),
            new("page_number", "numbering", AnalysisObjectKinds.PageNumber, RuleProperties.Numbering,
                false, "页码序列/前置罗马数字未读取"),

            // R4 Chinese abstract.
            new("abstract_cn.title", "font_cn", AnalysisObjectKinds.AbstractCnTitle, RuleProperties.FontEastAsia),
            new("abstract_cn.title", "size_pt", AnalysisObjectKinds.AbstractCnTitle, RuleProperties.FontSize),
            new("abstract_cn.title", "bold", AnalysisObjectKinds.AbstractCnTitle, RuleProperties.Bold),
            new("abstract_cn.title", "alignment", AnalysisObjectKinds.AbstractCnTitle, RuleProperties.Alignment),
            new("abstract_cn.title", "line_spacing", AnalysisObjectKinds.AbstractCnTitle, RuleProperties.LineSpacing),
            new("abstract_cn.title", "spacing.before_pt", AnalysisObjectKinds.AbstractCnTitle, RuleProperties.SpaceBefore),
            new("abstract_cn.title", "spacing.after_pt", AnalysisObjectKinds.AbstractCnTitle, RuleProperties.SpaceAfter),
            new("abstract_cn.body", "font_cn", AnalysisObjectKinds.AbstractCnBody, RuleProperties.FontEastAsia),
            new("abstract_cn.body", "size_pt", AnalysisObjectKinds.AbstractCnBody, RuleProperties.FontSize),
            new("abstract_cn.body", "line_spacing", AnalysisObjectKinds.AbstractCnBody, RuleProperties.LineSpacing),
            new("abstract_cn.body", "spacing.before_pt", AnalysisObjectKinds.AbstractCnBody, RuleProperties.SpaceBefore),
            new("abstract_cn.body", "spacing.after_pt", AnalysisObjectKinds.AbstractCnBody, RuleProperties.SpaceAfter),
            new("abstract_cn.body", "keyword_label_bold", AnalysisObjectKinds.AbstractCnBody, RuleProperties.KeywordBold),

            // R5 English abstract.
            new("abstract_en.title", "font_western", AnalysisObjectKinds.AbstractEnTitle, RuleProperties.FontLatin),
            new("abstract_en.title", "size_pt", AnalysisObjectKinds.AbstractEnTitle, RuleProperties.FontSize),
            new("abstract_en.title", "bold", AnalysisObjectKinds.AbstractEnTitle, RuleProperties.Bold),
            new("abstract_en.title", "alignment", AnalysisObjectKinds.AbstractEnTitle, RuleProperties.Alignment),
            new("abstract_en.title", "line_spacing", AnalysisObjectKinds.AbstractEnTitle, RuleProperties.LineSpacing),
            new("abstract_en.title", "spacing.before_pt", AnalysisObjectKinds.AbstractEnTitle, RuleProperties.SpaceBefore),
            new("abstract_en.title", "spacing.after_pt", AnalysisObjectKinds.AbstractEnTitle, RuleProperties.SpaceAfter),
            new("abstract_en.body", "font_western", AnalysisObjectKinds.AbstractEnBody, RuleProperties.FontLatin),
            new("abstract_en.body", "size_pt", AnalysisObjectKinds.AbstractEnBody, RuleProperties.FontSize),
            new("abstract_en.body", "line_spacing", AnalysisObjectKinds.AbstractEnBody, RuleProperties.LineSpacing),
            new("abstract_en.body", "spacing.before_pt", AnalysisObjectKinds.AbstractEnBody, RuleProperties.SpaceBefore),
            new("abstract_en.body", "spacing.after_pt", AnalysisObjectKinds.AbstractEnBody, RuleProperties.SpaceAfter),
            new("abstract_en.body", "keyword_label_bold", AnalysisObjectKinds.AbstractEnBody, RuleProperties.KeywordBold),

            // R6 table of contents: title is checked, entries wait for T6.
            new("toc.title", "font_cn", AnalysisObjectKinds.TocTitle, RuleProperties.FontEastAsia),
            new("toc.title", "size_pt", AnalysisObjectKinds.TocTitle, RuleProperties.FontSize),
            new("toc.title", "bold", AnalysisObjectKinds.TocTitle, RuleProperties.Bold),
            new("toc.title", "alignment", AnalysisObjectKinds.TocTitle, RuleProperties.Alignment),
            new("toc.title", "line_spacing", AnalysisObjectKinds.TocTitle, RuleProperties.LineSpacing),
            new("toc.title", "spacing.before_pt", AnalysisObjectKinds.TocTitle, RuleProperties.SpaceBefore),
            new("toc.title", "spacing.after_pt", AnalysisObjectKinds.TocTitle, RuleProperties.SpaceAfter),
            new("toc", "entries", AnalysisObjectKinds.TocEntry, RuleProperties.TocEntries,
                false, "目录条目检测对象（TOC 域样式 vs 渲染结果）待定（T6）",
                "章/一级/二级目录条目格式"),

            // R7-R10 headings.
            new("headings.chapter", "font_cn", AnalysisObjectKinds.HeadingChapter, RuleProperties.FontEastAsia),
            new("headings.chapter", "size_pt", AnalysisObjectKinds.HeadingChapter, RuleProperties.FontSize),
            new("headings.chapter", "bold", AnalysisObjectKinds.HeadingChapter, RuleProperties.Bold),
            new("headings.chapter", "alignment", AnalysisObjectKinds.HeadingChapter, RuleProperties.Alignment,
                false, "章标题对齐取决于编号模式（A2），本版不判定"),
            new("headings.chapter", "line_spacing", AnalysisObjectKinds.HeadingChapter, RuleProperties.LineSpacing),
            new("headings.chapter", "spacing.before_pt", AnalysisObjectKinds.HeadingChapter, RuleProperties.SpaceBefore),
            new("headings.chapter", "spacing.after_pt", AnalysisObjectKinds.HeadingChapter, RuleProperties.SpaceAfter),
            new("headings.level1", "font_cn", AnalysisObjectKinds.HeadingLevel1, RuleProperties.FontEastAsia),
            new("headings.level1", "size_pt", AnalysisObjectKinds.HeadingLevel1, RuleProperties.FontSize),
            new("headings.level1", "bold", AnalysisObjectKinds.HeadingLevel1, RuleProperties.Bold),
            new("headings.level1", "alignment", AnalysisObjectKinds.HeadingLevel1, RuleProperties.Alignment,
                false, "一级标题对齐取决于编号模式（A2），本版不判定"),
            new("headings.level1", "line_spacing", AnalysisObjectKinds.HeadingLevel1, RuleProperties.LineSpacing),
            new("headings.level1", "spacing.before_pt", AnalysisObjectKinds.HeadingLevel1, RuleProperties.SpaceBefore),
            new("headings.level1", "spacing.after_pt", AnalysisObjectKinds.HeadingLevel1, RuleProperties.SpaceAfter),
            new("headings.level2", "font_cn", AnalysisObjectKinds.HeadingLevel2, RuleProperties.FontEastAsia),
            new("headings.level2", "size_pt", AnalysisObjectKinds.HeadingLevel2, RuleProperties.FontSize),
            new("headings.level2", "bold", AnalysisObjectKinds.HeadingLevel2, RuleProperties.Bold),
            new("headings.level2", "alignment", AnalysisObjectKinds.HeadingLevel2, RuleProperties.Alignment),
            new("headings.level2", "line_spacing", AnalysisObjectKinds.HeadingLevel2, RuleProperties.LineSpacing),
            new("headings.level2", "spacing.before_pt", AnalysisObjectKinds.HeadingLevel2, RuleProperties.SpaceBefore),
            new("headings.level2", "spacing.after_pt", AnalysisObjectKinds.HeadingLevel2, RuleProperties.SpaceAfter),
            new("headings.level2", "indent_chars_mode2", AnalysisObjectKinds.HeadingLevel2, RuleProperties.FirstLineIndent,
                false, "模式二左缩进需先判定编号模式（A2）"),
            new("headings.level3", "font_cn", AnalysisObjectKinds.HeadingLevel3, RuleProperties.FontEastAsia),
            new("headings.level3", "size_pt", AnalysisObjectKinds.HeadingLevel3, RuleProperties.FontSize),
            new("headings.level3", "bold", AnalysisObjectKinds.HeadingLevel3, RuleProperties.Bold),
            new("headings.level3", "alignment", AnalysisObjectKinds.HeadingLevel3, RuleProperties.Alignment),
            new("headings.level3", "line_spacing", AnalysisObjectKinds.HeadingLevel3, RuleProperties.LineSpacing),
            new("headings.level3", "spacing.before_pt", AnalysisObjectKinds.HeadingLevel3, RuleProperties.SpaceBefore),
            new("headings.level3", "spacing.after_pt", AnalysisObjectKinds.HeadingLevel3, RuleProperties.SpaceAfter),
            new("headings.level3", "indent_chars_mode2", AnalysisObjectKinds.HeadingLevel3, RuleProperties.FirstLineIndent,
                false, "模式二左缩进需先判定编号模式（A2）"),

            // R11 body text.
            new("body_text", "font.cn", AnalysisObjectKinds.BodyText, RuleProperties.FontEastAsia),
            new("body_text", "font.western", AnalysisObjectKinds.BodyText, RuleProperties.FontLatin),
            new("body_text", "size_pt", AnalysisObjectKinds.BodyText, RuleProperties.FontSize),
            new("body_text", "alignment", AnalysisObjectKinds.BodyText, RuleProperties.Alignment),
            new("body_text", "first_line_indent_chars", AnalysisObjectKinds.BodyText, RuleProperties.FirstLineIndent),
            new("body_text", "line_spacing", AnalysisObjectKinds.BodyText, RuleProperties.LineSpacing),
            new("body_text", "spacing.before_pt", AnalysisObjectKinds.BodyText, RuleProperties.SpaceBefore),
            new("body_text", "spacing.after_pt", AnalysisObjectKinds.BodyText, RuleProperties.SpaceAfter),

            // R12/R13 are owned by the figure/table work (#92); this analyser must not duplicate it.
            new("figures_and_tables.caption", string.Empty, AnalysisObjectKinds.Caption, RuleProperties.Caption,
                false, "图表与题注规则由 #92 的 FigureTable* 实现，本分析器不重复检查",
                "图题/表题格式（宋体 10.5pt、居中、单倍、段前6/段后12或6）"),
            new("figures_and_tables.equation", string.Empty, AnalysisObjectKinds.Equation, RuleProperties.Equation,
                false, "公式对象识别与编号右对齐由 #92/后续工作实现",
                "公式编号 Times New Roman 10.5pt、右对齐"),

            // R14 references body.
            new("references", "font.cn", AnalysisObjectKinds.ReferencesBody, RuleProperties.FontEastAsia),
            new("references", "font.western", AnalysisObjectKinds.ReferencesBody, RuleProperties.FontLatin),
            new("references", "size_pt", AnalysisObjectKinds.ReferencesBody, RuleProperties.FontSize),
            new("references", "line_spacing", AnalysisObjectKinds.ReferencesBody, RuleProperties.LineSpacing),
            new("references", "spacing.before_pt", AnalysisObjectKinds.ReferencesBody, RuleProperties.SpaceBefore),
            new("references", "spacing.after_pt", AnalysisObjectKinds.ReferencesBody, RuleProperties.SpaceAfter),
            new("references", "heading", AnalysisObjectKinds.ReferencesBody, RuleProperties.HeadingReuse,
                false, "参考文献标题复用 R7 章标题规则，见 R7 覆盖行",
                "同章标题"),

            // R15 appendix body.
            new("appendix", "font.cn", AnalysisObjectKinds.AppendixBody, RuleProperties.FontEastAsia),
            new("appendix", "font.western", AnalysisObjectKinds.AppendixBody, RuleProperties.FontLatin),
            new("appendix", "size_pt", AnalysisObjectKinds.AppendixBody, RuleProperties.FontSize),
            new("appendix", "alignment", AnalysisObjectKinds.AppendixBody, RuleProperties.Alignment),
            new("appendix", "first_line_indent_chars", AnalysisObjectKinds.AppendixBody, RuleProperties.FirstLineIndent),
            new("appendix", "line_spacing", AnalysisObjectKinds.AppendixBody, RuleProperties.LineSpacing),
            new("appendix", "spacing.before_pt", AnalysisObjectKinds.AppendixBody, RuleProperties.SpaceBefore),
            new("appendix", "spacing.after_pt", AnalysisObjectKinds.AppendixBody, RuleProperties.SpaceAfter),
            new("appendix", "heading", AnalysisObjectKinds.AppendixBody, RuleProperties.HeadingReuse,
                false, "附录标题复用 R7 章标题规则，见 R7 覆盖行",
                "同章标题"),

            // R16 symbol table / CV: recognition is not implemented yet.
            new("symbol_table_and_cv", "font.cn", AnalysisObjectKinds.SymbolTable, RuleProperties.FontEastAsia,
                false, "符号注释表/个人简历的对象识别未实现"),
            new("symbol_table_and_cv", "font.western", AnalysisObjectKinds.SymbolTable, RuleProperties.FontLatin,
                false, "符号注释表/个人简历的对象识别未实现"),
            new("symbol_table_and_cv", "size_pt", AnalysisObjectKinds.SymbolTable, RuleProperties.FontSize,
                false, "符号注释表/个人简历的对象识别未实现"),
            new("symbol_table_and_cv", "line_spacing", AnalysisObjectKinds.SymbolTable, RuleProperties.LineSpacing,
                false, "符号注释表/个人简历的对象识别未实现"),
            new("symbol_table_and_cv", "spacing.before_pt", AnalysisObjectKinds.SymbolTable, RuleProperties.SpaceBefore,
                false, "符号注释表/个人简历的对象识别未实现"),
            new("symbol_table_and_cv", "spacing.after_pt", AnalysisObjectKinds.SymbolTable, RuleProperties.SpaceAfter,
                false, "符号注释表/个人简历的对象识别未实现"),
            new("symbol_table_and_cv", "heading", AnalysisObjectKinds.SymbolTable, RuleProperties.HeadingReuse,
                false, "符号注释表/个人简历标题复用 R7 章标题规则，但对象识别尚未实现",
                "同章标题"),

            // R17 acknowledgements body.
            new("acknowledgement", "font_cn", AnalysisObjectKinds.AcknowledgementBody, RuleProperties.FontEastAsia),
            new("acknowledgement", "size_pt", AnalysisObjectKinds.AcknowledgementBody, RuleProperties.FontSize),
            new("acknowledgement", "line_spacing", AnalysisObjectKinds.AcknowledgementBody, RuleProperties.LineSpacing),
            new("acknowledgement", "spacing.before_pt", AnalysisObjectKinds.AcknowledgementBody, RuleProperties.SpaceBefore),
            new("acknowledgement", "spacing.after_pt", AnalysisObjectKinds.AcknowledgementBody, RuleProperties.SpaceAfter),
            new("acknowledgement", "heading", AnalysisObjectKinds.AcknowledgementBody, RuleProperties.HeadingReuse,
                false, "致谢标题复用 R7 章标题规则，见 R7 覆盖行",
                "同章标题"),

            // R18 cover: recorded but explicitly skipped.
            new("cover", string.Empty, AnalysisObjectKinds.Cover, RuleProperties.Cover,
                false, "封面为学校固定版式，v1 不检查（RFC-6）",
                "封面/题名页固定版式")
        };
        return slots.AsReadOnly();
    }
}
