using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class ThesisStructureClassifierTests
{
    private readonly ThesisStructureClassifier _classifier = new();

    [Fact]
    public void ClassifiesUnstyledChineseAndEnglishThesisStructureConservatively()
    {
        var blocks = new[]
        {
            B("title", 0, "南开大学研究生学位论文（合成测试样稿）"),
            B("cn-abstract-title", 1, "中文摘要"),
            B("cn-abstract", 2, "本研究观察高校图书馆学习空间的使用变化，并分析不同区域的服务需求。"),
            B("cn-keywords", 3, "关键词：空间治理；学习空间；公共服务"),
            B("en-abstract-title", 4, "英文摘要"),
            B("en-abstract", 5, "This synthetic study examines learning-space use and service needs across university library zones."),
            B("en-keywords", 6, "Keywords: learning spaces; public service"),
            B("toc-title", 7, "目录"),
            B("toc-chapter", 8, "第一章 绪论 ........ 1"),
            B("toc-section", 9, "1.1 研究背景 ........ 2"),
            B("toc-method", 10, "Chapter 2 Methods ........ 8"),
            B("toc-refs", 11, "参考文献 ........ 25"),
            B("body-chapter", 12, "第一章 绪论"),
            B("body-opening", 13, "随着移动设备普及，学生会在教学楼、宿舍和图书馆之间灵活安排学习活动，空间选择受到课程、设备和协作需求影响。"),
            B("body-section", 14, "1.1 研究背景"),
            B("body-reference-mention", 15, "本文在正文中讨论参考文献的组织方式，但这句话不是参考文献章节标题。"),
            B("refs-title", 16, "参考文献"),
            B("refs-entry", 17, "[1] Wang, Li. A synthetic study of library learning spaces. 2025."),
            B("appendix-title", 18, "附录 A 调查工具"),
            B("appendix-text", 19, "本附录列出合成问卷的字段。"),
            B("ack-title", 20, "致谢"),
            B("ack-text", 21, "感谢参与合成测试的研究人员。")
        };

        var result = _classifier.Classify(blocks);
        var byId = result.Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(byId["title"], ThesisSectionKind.FrontMatter, ThesisBlockKind.Paragraph);
        AssertSectionAndKind(byId["cn-abstract-title"], ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["cn-abstract"], ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Paragraph);
        AssertSectionAndKind(byId["cn-keywords"], ThesisSectionKind.ChineseAbstract, ThesisBlockKind.Keywords);
        AssertSectionAndKind(byId["en-abstract-title"], ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["en-abstract"], ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Paragraph);
        AssertSectionAndKind(byId["en-keywords"], ThesisSectionKind.EnglishAbstract, ThesisBlockKind.Keywords);
        AssertSectionAndKind(byId["toc-title"], ThesisSectionKind.Contents, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["toc-chapter"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["toc-section"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["toc-method"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["toc-refs"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["body-chapter"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["body-opening"], ThesisSectionKind.Body, ThesisBlockKind.Paragraph);
        AssertSectionAndKind(byId["body-section"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["body-reference-mention"], ThesisSectionKind.Body, ThesisBlockKind.Paragraph);
        AssertSectionAndKind(byId["refs-title"], ThesisSectionKind.References, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["refs-entry"], ThesisSectionKind.References, ThesisBlockKind.Paragraph);
        AssertSectionAndKind(byId["appendix-title"], ThesisSectionKind.Appendix, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["ack-title"], ThesisSectionKind.Acknowledgments, ThesisBlockKind.Heading);

        Assert.False(byId["body-chapter"].RequiresReview);
        Assert.False(byId["body-opening"].RequiresReview);
        Assert.False(byId["body-reference-mention"].RequiresReview);
        Assert.True(byId["body-chapter"].Confidence >= 0.92);
        Assert.InRange(byId["body-chapter"].Confidence, 0.0, 1.0);
        Assert.DoesNotContain(result.Warnings, warning => warning.Contains("正文起始标题"));
    }

    [Fact]
    public void KeepsRepeatedTocLabelsInContentsAndRecognizesActualBodyAndReferences()
    {
        var blocks = new[]
        {
            B("contents", 0, "目录"),
            B("toc-chapter", 1, "第一章 绪论 …… 1"),
            B("toc-references", 2, "参考文献 …… 20"),
            B("toc-appendix", 3, "附录 A 工具 …… 21"),
            B("actual-chapter", 4, "第一章 绪论"),
            B("body", 5, "本研究采用合成数据观察学习空间变化，下面的分析仅用于检验结构识别和修复流程。"),
            B("body-mention", 6, "参考文献部分将在后文讨论；这不是章节标题，也不应结束正文范围。"),
            B("actual-references", 7, "参考文献"),
            B("reference", 8, "[1] Example, A. A synthetic reference entry. 2026.")
        };

        var byId = _classifier.Classify(blocks).Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(byId["toc-chapter"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["toc-references"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["toc-appendix"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["actual-chapter"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["body-mention"], ThesisSectionKind.Body, ThesisBlockKind.Paragraph);
        AssertSectionAndKind(byId["actual-references"], ThesisSectionKind.References, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["reference"], ThesisSectionKind.References, ThesisBlockKind.Paragraph);
    }

    [Fact]
    public void DoesNotForceUnmarkedLongTextOrSentenceLikeNumberedTextIntoBody()
    {
        var blocks = new[]
        {
            B("long-unmarked", 0, "This long synthetic paragraph discusses chapter one and the phrase 第一章 绪论 in passing, but contains no standalone structural heading or reliable boundary."),
            B("numbered-sentence", 1, "1 Introduction explains how the synthetic observation process works, and this is a sentence rather than a short heading."),
            B("long-unmarked-cn", 2, "第一章 绪论在本研究中只是句子开头，后面继续描述研究对象、方法和结果，没有标题边界。")
        };

        var result = _classifier.Classify(blocks);

        Assert.All(result.Blocks, decision =>
        {
            Assert.Equal(ThesisSectionKind.Unknown, decision.Section);
            Assert.Equal(ThesisBlockKind.Unknown, decision.Kind);
            Assert.True(decision.RequiresReview);
        });
        Assert.Contains(result.Warnings, warning => warning.Contains("没有把全文默认分类为正文"));
    }

    [Fact]
    public void MarksUnsupportedInlineContentAsUnknownForReviewWithoutLosingSectionContext()
    {
        var blocks = new[]
        {
            B("chapter", 0, "第一章 绪论"),
            B("paragraph", 1, "本段提供足够长的合成正文内容，用于建立正文范围并验证后续结构识别。"),
            B("unsupported", 2, "包含公式或复杂域的文本快照", unsupportedInline: true),
            B("next-paragraph", 3, "后续的合成正文仍然可以沿用已经确认的正文范围。")
        };

        var result = _classifier.Classify(blocks);
        var byId = result.Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(byId["chapter"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        Assert.Equal(ThesisSectionKind.Body, byId["unsupported"].Section);
        Assert.Equal(ThesisBlockKind.Unknown, byId["unsupported"].Kind);
        Assert.True(byId["unsupported"].RequiresReview);
        Assert.Equal(ThesisBlockKind.Paragraph, byId["next-paragraph"].Kind);
        Assert.True(byId["next-paragraph"].RequiresReview);
        Assert.Contains(result.Warnings, warning => warning.Contains("未支持的内联结构"));
    }

    [Fact]
    public void ExcludesTableCellsAndCaptionsFromBodyProseClassification()
    {
        var blocks = new[]
        {
            B("chapter", 0, "第一章 绪论"),
            B("prose", 1, "正文说明在表格之前出现，满足形成明确正文范围所需的结构条件。"),
            B("table-cell", 2, "虚构区域", tableCell: true),
            B("explicit-table-caption", 3, "表 1.1 样本字段", caption: CaptionKind.Table),
            B("text-figure-caption", 4, "图 1.1 合成区域示意图"),
            B("after-caption", 5, "图表对象之后仍是正常的合成正文段落，用于确认分类器继续保留正文范围。")
        };

        var byId = _classifier.Classify(blocks).Blocks.ToDictionary(decision => decision.BlockId);

        Assert.Equal(ThesisBlockKind.TableCell, byId["table-cell"].Kind);
        Assert.Equal(ThesisBlockKind.TableCaption, byId["explicit-table-caption"].Kind);
        Assert.Equal(ThesisBlockKind.FigureCaption, byId["text-figure-caption"].Kind);
        Assert.True(byId["text-figure-caption"].RequiresReview);
        Assert.NotEqual(ThesisBlockKind.Paragraph, byId["table-cell"].Kind);
        Assert.NotEqual(ThesisBlockKind.Paragraph, byId["explicit-table-caption"].Kind);
        Assert.NotEqual(ThesisBlockKind.Paragraph, byId["text-figure-caption"].Kind);
        Assert.Equal(ThesisBlockKind.Paragraph, byId["after-caption"].Kind);
    }

    [Fact]
    public void UsesWordOutlineMetadataAsStrongHeadingEvidence()
    {
        var result = _classifier.Classify(new[]
        {
            new ThesisStructureBlock("styled", 0, "第一章 研究设计", HasHeadingStyle: true, OutlineLevel: 1),
            B("text", 1, "此处为合成的章节正文，用于验证标题元数据能够确定章节范围。")
        });

        Assert.Equal(ThesisBlockKind.Heading, result.Blocks[0].Kind);
        Assert.Equal(ThesisSectionKind.Body, result.Blocks[0].Section);
        Assert.Equal(2, result.Blocks[0].HeadingLevel);
        Assert.False(result.Blocks[0].RequiresReview);
        Assert.Equal(ThesisBlockKind.Paragraph, result.Blocks[1].Kind);
    }

    [Fact]
    public void PromotesUnstyledChapterBodyOnlyWhenAnIndependentClosingSectionConfirmsTheScope()
    {
        var closed = _classifier.Classify(new[]
        {
            B("chapter", 0, "Chapter 1 Introduction"),
            B("opening", 1, "This synthetic paragraph describes the study scope and provides ordinary prose after the chapter heading."),
            B("body-section", 2, "1.1 Background"),
            B("body-detail", 3, "The following synthetic paragraph adds ordinary prose within the candidate body range."),
            B("references", 4, "References"),
            B("reference-entry", 5, "[1] Example, A. Synthetic reference. 2026.")
        }).Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(closed["chapter"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        AssertSectionAndKind(closed["opening"], ThesisSectionKind.Body, ThesisBlockKind.Paragraph);
        Assert.True(closed["chapter"].Confidence >= 0.92);
        Assert.True(closed["opening"].Confidence >= 0.92);
        Assert.False(closed["chapter"].RequiresReview);
        Assert.False(closed["opening"].RequiresReview);
        Assert.False(closed["body-detail"].RequiresReview);
        AssertSectionAndKind(closed["references"], ThesisSectionKind.References, ThesisBlockKind.Heading);

        var openEnded = _classifier.Classify(new[]
        {
            B("chapter", 0, "第一章 绪论"),
            B("opening", 1, "本合成段落描述研究背景、研究目标和分析范围，构成章节标题之后的普通正文。"),
            B("subheading", 2, "1.1 研究背景"),
            B("body-detail", 3, "这里还有一段普通正文，但文档中没有独立参考文献、附录或致谢标题。")
        }).Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(openEnded["chapter"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        Assert.True(openEnded["chapter"].Confidence < 0.90);
        Assert.True(openEnded["chapter"].RequiresReview);
        Assert.True(openEnded["opening"].RequiresReview);
        Assert.True(openEnded["subheading"].RequiresReview);
        Assert.True(openEnded["body-detail"].RequiresReview);
    }

    [Fact]
    public void PromotesCoherentMultilevelUnstyledNumberingWhenBodyTextAndClosingBoundaryAgree()
    {
        var byId = _classifier.Classify(new[]
        {
            B("first-subheading", 0, "1.1 Background"),
            B("first-prose", 1, "This synthetic paragraph describes the background using ordinary prose and no Word heading styles."),
            B("next-subheading", 2, "1.2 Methods"),
            B("second-prose", 3, "This synthetic paragraph describes the method and supplies further prose in the same numbered section."),
            B("references", 4, "References")
        }).Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(byId["first-subheading"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        Assert.True(byId["first-subheading"].Confidence >= 0.92);
        Assert.False(byId["first-subheading"].RequiresReview);
        Assert.False(byId["first-prose"].RequiresReview);
        Assert.False(byId["second-prose"].RequiresReview);
    }

    [Fact]
    public void AcceptsUnnumberedSemanticHeadingsOnlyWithSupportingStructure()
    {
        var closed = _classifier.Classify(new[]
        {
            B("introduction", 0, "绪论"),
            B("intro-prose", 1, "本合成段落介绍研究问题、研究目标以及论文的组织结构，作为正文开篇内容。"),
            B("methods", 2, "研究方法"),
            B("methods-prose", 3, "此合成段落说明资料收集步骤和分析方法，并构成普通正文。"),
            B("results", 4, "Results"),
            B("results-prose", 5, "The synthetic results paragraph describes aggregate observations without using real student data."),
            B("discussion", 6, "讨论"),
            B("discussion-prose", 7, "这里是对合成观察结果的讨论内容，属于正文段落。"),
            B("conclusions", 8, "Conclusions"),
            B("conclusion-prose", 9, "This synthetic conclusion summarizes the findings and limitations of the demonstration."),
            B("references", 10, "参考文献")
        }).Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(closed["introduction"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        Assert.True(closed["introduction"].Confidence >= 0.92);
        Assert.False(closed["introduction"].RequiresReview);
        AssertSectionAndKind(closed["methods"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        AssertSectionAndKind(closed["results"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        Assert.False(closed["conclusion-prose"].RequiresReview);

        var unclosed = _classifier.Classify(new[]
        {
            B("introduction", 0, "Introduction"),
            B("intro-prose", 1, "This synthetic paragraph introduces the subject and establishes ordinary body prose."),
            B("methods", 2, "Methods"),
            B("methods-prose", 3, "This synthetic paragraph describes the method, but no independent closing section follows."),
            B("results", 4, "Results"),
            B("results-prose", 5, "This synthetic result paragraph continues the candidate document body.")
        }).Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(unclosed["introduction"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        Assert.True(unclosed["introduction"].Confidence < 0.90);
        Assert.True(unclosed["introduction"].RequiresReview);
        Assert.True(unclosed["intro-prose"].RequiresReview);
        Assert.True(unclosed["results-prose"].RequiresReview);
    }

    [Fact]
    public void ASingleUnnumberedIntroductionWordDoesNotClaimTheWholeDocumentAsBody()
    {
        var byId = _classifier.Classify(new[]
        {
            B("introduction", 0, "Introduction"),
            B("prose", 1, "This synthetic paragraph follows the standalone title, but there are no other body headings."),
            B("references", 2, "References")
        }).Blocks.ToDictionary(decision => decision.BlockId);

        Assert.Equal(ThesisSectionKind.Unknown, byId["introduction"].Section);
        Assert.Equal(ThesisBlockKind.Heading, byId["introduction"].Kind);
        Assert.True(byId["introduction"].RequiresReview);
        Assert.Equal(ThesisSectionKind.Unknown, byId["prose"].Section);
        Assert.Equal(ThesisBlockKind.Unknown, byId["prose"].Kind);
        Assert.True(byId["prose"].RequiresReview);
    }

    [Fact]
    public void DoesNotPromoteAnUnpaginatedTocEntryByBorrowingProseAcrossTheNextChapterRoot()
    {
        var byId = _classifier.Classify(new[]
        {
            B("toc-title", 0, "目录"),
            B("toc-chapter-one", 1, "第一章 绪论"),
            B("toc-chapter-two", 2, "第二章 方法"),
            B("actual-chapter-one", 3, "第一章 绪论"),
            B("actual-opening", 4, "随着移动设备普及，高校学生会在教学楼、宿舍和图书馆之间灵活安排学习活动，研究将分析空间需求变化。"),
            B("references", 5, "参考文献"),
            B("reference-entry", 6, "[1] Example. A synthetic reference. 2026.")
        }).Blocks.ToDictionary(decision => decision.BlockId);

        AssertSectionAndKind(byId["toc-title"], ThesisSectionKind.Contents, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["toc-chapter-one"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["toc-chapter-two"], ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry);
        AssertSectionAndKind(byId["actual-chapter-one"], ThesisSectionKind.Body, ThesisBlockKind.Heading);
        AssertSectionAndKind(byId["actual-opening"], ThesisSectionKind.Body, ThesisBlockKind.Paragraph);
        Assert.False(byId["actual-chapter-one"].RequiresReview);
        Assert.False(byId["actual-opening"].RequiresReview);
    }

    [Fact]
    public void HandlesTenThousandHeadingOnlyBlocksWithoutInventingABodyRegion()
    {
        var headings = Enumerable.Range(0, 10_000)
            .Select(index => B($"heading-{index}", index, $"第{index + 1}章 合成标题"))
            .ToArray();

        var result = _classifier.Classify(headings);

        Assert.Equal(10_000, result.Blocks.Count);
        Assert.DoesNotContain(result.Blocks, decision => decision.Section == ThesisSectionKind.Body);
        Assert.All(result.Blocks, decision => Assert.True(decision.RequiresReview));
        Assert.Contains(result.Warnings, warning => warning.Contains("没有把全文默认分类为正文"));
    }

    private static ThesisStructureBlock B(
        string id,
        int paragraphIndex,
        string text,
        bool tableCell = false,
        CaptionKind? caption = null,
        bool unsupportedInline = false) =>
        new(id, paragraphIndex, text, IsTableCell: tableCell, Caption: caption,
            HasUnsupportedInlineContent: unsupportedInline);

    private static void AssertSectionAndKind(
        ThesisStructureDecision decision,
        ThesisSectionKind section,
        ThesisBlockKind kind)
    {
        Assert.Equal(section, decision.Section);
        Assert.Equal(kind, decision.Kind);
    }
}
