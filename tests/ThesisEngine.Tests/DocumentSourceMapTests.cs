using System.Xml.Linq;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DocumentSourceMapTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    [Fact]
    public void UniqueTextRelocatesAfterParagraphInsertionWithoutUsingOldOrdinal()
    {
        var before = Map(P("摘要"), P("  保留原样的正文 English  "), P("参考文献"));
        var after = Map(P("封面"), P("摘要"), P("  保留原样的正文 English  "), P("参考文献"));

        var resolution = after.Resolve(before.Blocks[1]);

        Assert.Equal(DocumentAnchorResolutionStatus.Relocated, resolution.Status);
        Assert.Equal(2, resolution.Block!.ParagraphIndex);
        Assert.Equal("  保留原样的正文 English  ", resolution.Block.ExactText);
        Assert.Equal(before.Blocks[1].BlockId, resolution.Block.BlockId);
        Assert.Equal(before.Blocks[1].TextHash, resolution.Block.TextHash);
    }

    [Fact]
    public void WordParagraphIdWinsWhenSameTextExistsElsewhereButEditedIdentityIsRejected()
    {
        var before = Map(P("相同句子", "001122AA"), P("相同句子", "001122BB"));
        var moved = Map(P("插入段落"), P("相同句子", "001122BB"), P("相同句子", "001122AA"));
        var resolution = moved.Resolve(before.Blocks[0]);
        Assert.Equal(DocumentAnchorResolutionStatus.Relocated, resolution.Status);
        Assert.Equal(2, resolution.Block!.ParagraphIndex);

        var edited = Map(P("被用户改过的句子", "001122AA"), P("相同句子", "001122BB"));
        var refused = edited.Resolve(before.Blocks[0]);
        Assert.Equal(DocumentAnchorResolutionStatus.TextChanged, refused.Status);
        Assert.Null(refused.Block);
    }

    [Fact]
    public void FormattingAndRunSplitsKeepBlockIdentityAndSpanIsProjectedToCurrentRuns()
    {
        var before = Map(P("中文 English 样例"));
        var currentParagraph = new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "both"))),
            Run("中文 ", "Arial"), Run("Eng", "Times New Roman"), Run("lish 样例", "宋体"));
        var after = Map(currentParagraph);

        Assert.Equal(before.Blocks[0].BlockId, after.Blocks[0].BlockId);
        var result = after.ResolveSpan(before.Blocks[0], 3, 7, "English");

        Assert.True(result.IsResolved);
        Assert.Equal(new[] { "Eng", "lish" }, result.Segments.Select(segment => segment.ExactText));
        Assert.Equal(new[] { 1, 2 }, result.Segments.Select(segment => segment.RunIndex));
        Assert.Equal("word/document.xml#paragraph[0]/run[1]", result.Segments[0].LegacyLocation);
        Assert.Equal(0, result.Segments[1].RunUtf16Offset);
    }

    [Fact]
    public void IdenticalShortParagraphsWithIdenticalContextsRemainAmbiguous()
    {
        var before = Map(P("背景"), P("重复"), P("边界"), P("背景"), P("重复"), P("边界"));
        var after = Map(P("背景"), P("重复"), P("边界"), P("背景"), P("重复"), P("边界"));

        var result = after.Resolve(before.Blocks[1]);

        Assert.Equal(DocumentAnchorResolutionStatus.Ambiguous, result.Status);
        Assert.Null(result.Block);
    }

    [Fact]
    public void RepeatedParagraphCanResolveOnlyWithUniqueUnchangedContextAndCount()
    {
        var before = Map(P("第一节"), P("重复"), P("第一节结束"), P("第二节"), P("重复"), P("第二节结束"));
        var after = Map(P("新增封面"), P("第一节"), P("重复"), P("第一节结束"), P("第二节"), P("重复"), P("第二节结束"));

        var result = after.Resolve(before.Blocks[4]);

        Assert.Equal(DocumentAnchorResolutionStatus.Relocated, result.Status);
        Assert.Equal(5, result.Block!.ParagraphIndex);
    }

    [Fact]
    public void InsertingIdenticalParagraphCannotStealOldOrdinalOrBoundaryContext()
    {
        var before = Map(P("重复"), P("重复"));
        var after = Map(P("重复"), P("重复"), P("重复"));

        var result = after.Resolve(before.Blocks[0]);

        Assert.Equal(DocumentAnchorResolutionStatus.Ambiguous, result.Status);
        Assert.Null(result.Block);
    }

    [Fact]
    public void NeighborInsertionRefusesRepeatedTextWhenItsOldContextNoLongerMatches()
    {
        var before = Map(P("第一节"), P("重复"), P("第一节结束"), P("第二节"), P("重复"), P("第二节结束"));
        var after = Map(P("第一节"), P("新插入解释"), P("重复"), P("第一节结束"), P("第二节"), P("重复"), P("第二节结束"));

        var result = after.Resolve(before.Blocks[1]);

        Assert.Equal(DocumentAnchorResolutionStatus.Ambiguous, result.Status);
        Assert.Null(result.Block);
    }

    [Fact]
    public void DuplicateWordIdsDoNotSilentlySelectOneParagraph()
    {
        var before = Map(P("正文", "001122AA"));
        var after = Map(P("正文", "001122AA"), P("正文", "001122AA"));

        Assert.Equal(DocumentAnchorResolutionStatus.Ambiguous, after.Resolve(before.Blocks[0]).Status);
        Assert.All(after.Blocks, block => Assert.False(block.HasUniqueParagraphId));
        Assert.Equal(2, after.Blocks.Select(block => block.BlockId).Distinct().Count());
    }

    [Fact]
    public void ChangedTextIsNotFuzzilyMatchedAndPartIdentityCannotBeSubstituted()
    {
        var before = Map(P("本段精确文字。"));
        var after = Map(P("本段精确文字!"));
        Assert.Equal(DocumentAnchorResolutionStatus.NotFound, after.Resolve(before.Blocks[0]).Status);
        var header = DocumentSourceMap.Create("word/header1.xml", new[] { P("本段精确文字。") });
        Assert.Equal(DocumentAnchorResolutionStatus.NotFound, header.Resolve(before.Blocks[0]).Status);
    }

    [Fact]
    public void PathsPreserveTableStructureAndUseSiblingNameOrdinals()
    {
        var bodyParagraph = P("正文");
        var firstCellParagraph = P("单元格一");
        var secondCellParagraph = P("单元格二");
        var finalParagraph = P("表后正文");
        var document = new XDocument(new XElement(W + "document", new XElement(W + "body", bodyParagraph,
            new XElement(W + "tbl", new XElement(W + "tr", new XElement(W + "tc", firstCellParagraph), new XElement(W + "tc", secondCellParagraph))),
            finalParagraph)));
        var paragraphs = document.Descendants(W + "p").ToArray();

        var map = DocumentSourceMap.Create("word/document.xml", paragraphs);

        Assert.Equal("/w:document[0]/w:body[0]/w:p[0]", map.Blocks[0].StructuralPath);
        Assert.Equal("/w:document[0]/w:body[0]/w:tbl[0]/w:tr[0]/w:tc[1]/w:p[0]", map.Blocks[2].StructuralPath);
        Assert.Equal("/w:document[0]/w:body[0]/w:p[1]", map.Blocks[3].StructuralPath);
        Assert.Equal("word/document.xml#paragraph[2]/run[0]", map.Blocks[2].Runs[0].LegacyLocation);
    }

    [Fact]
    public void ChineseAndEmojiSpansUseUtf16WithoutSplittingSurrogatePairs()
    {
        var before = Map(P("甲😀乙 English"));
        var after = Map(new XElement(W + "p", Run("甲"), Run("😀"), Run("乙 English")));

        var emoji = after.ResolveSpan(before.Blocks[0], 1, 2, "😀");
        Assert.True(emoji.IsResolved);
        var segment = Assert.Single(emoji.Segments);
        Assert.Equal(1, segment.ParagraphUtf16Offset);
        Assert.Equal(2, segment.Utf16Length);
        Assert.Equal(1, segment.RunIndex);
        Assert.Equal(DocumentSpanResolutionStatus.InvalidSpan, after.ResolveSpan(before.Blocks[0], 2, 1, "\uDE00").Status);
        Assert.Equal(DocumentSpanResolutionStatus.InvalidSpan, after.ResolveSpan(before.Blocks[0], 1, 1, "\uD83D").Status);
        Assert.Equal(DocumentSpanResolutionStatus.TextMismatch, after.ResolveSpan(before.Blocks[0], 1, 2, "xx").Status);
    }

    [Theory]
    [InlineData("tab")]
    [InlineData("br")]
    [InlineData("cr")]
    public void NontextWordCharactersNeverBecomeWritableXmlTextSpans(string element)
    {
        var map = Map(new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "甲"), new XElement(W + element), new XElement(W + "t", "乙"))));
        var block = Assert.Single(map.Blocks);

        Assert.Equal("甲乙", block.ExactText);
        Assert.False(block.HasCompleteTextMapping);
        Assert.Contains(block.UnsupportedReasons, reason => reason.Contains("w:" + element, StringComparison.Ordinal));
        Assert.Equal(DocumentSpanResolutionStatus.UnsupportedStructure, map.ResolveSpan(block, 0, 2, "甲乙").Status);
    }

    [Fact]
    public void HyperlinkTextIsPreservedButCannotBeMistakenForDirectRunOffsets()
    {
        var paragraph = new XElement(W + "p", Run("前文"), new XElement(W + "hyperlink", Run("链接")), Run("后文"));
        var map = Map(paragraph);
        var block = Assert.Single(map.Blocks);

        Assert.Equal("前文链接后文", block.ExactText);
        Assert.Equal(4, block.Runs[1].Utf16Offset);
        Assert.False(block.HasCompleteTextMapping);
        Assert.Equal(DocumentSpanResolutionStatus.UnsupportedStructure, map.ResolveSpan(block, 4, 2, "后文").Status);
    }

    [Fact]
    public void TextlessCommentRunsDoNotRenumberTextRunsAndFormattingChangesKeepId()
    {
        var before = Map(P("甲乙"));
        var after = Map(new XElement(W + "p", new XElement(W + "bookmarkStart", new XAttribute(W + "id", "1")),
            new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", "2"))), Run("甲乙", "宋体")));

        Assert.Equal(before.Blocks[0].BlockId, after.Blocks[0].BlockId);
        Assert.True(after.Blocks[0].HasCompleteTextMapping);
        Assert.Equal(0, Assert.Single(after.Blocks[0].Runs).RunIndex);
    }

    [Fact]
    public void NearestNonblankContextsSkipBlankParagraphsWithoutNormalizingSourceText()
    {
        var map = Map(P("前文"), P("  "), P("\t"), P("正文"), P(""), P("后文"));

        Assert.Equal(map.Blocks[0].TextHash, map.Blocks[3].PreviousTextHash);
        Assert.Equal(map.Blocks[5].TextHash, map.Blocks[3].NextTextHash);
        Assert.Equal("  ", map.Blocks[1].ExactText);
        Assert.NotEqual(Map(P("正文 ")).Blocks[0].TextHash, map.Blocks[3].TextHash);
    }

    [Fact]
    public void DuplicateInputsAndTextBoxParagraphsAreRejectedWithoutReindexing()
    {
        var paragraph = P("原文");
        Assert.Throws<ArgumentException>(() => DocumentSourceMap.Create("word/document.xml", new[] { paragraph, paragraph }));
        var textBoxParagraph = P("文本框");
        _ = new XElement(W + "txbxContent", textBoxParagraph);
        Assert.Throws<ArgumentException>(() => DocumentSourceMap.Create("word/document.xml", new[] { textBoxParagraph }));
    }

    [Fact]
    public void LookupUsesSnapshotOwnedBlockIdAndRejectsInventedTextMetadata()
    {
        var map = Map(P("原文"));
        var block = Assert.Single(map.Blocks);
        Assert.True(map.TryGetBlock(block.BlockId, out var indexed));
        Assert.Same(block, indexed);
        Assert.False(map.TryGetBlock("invented", out _));
        Assert.Equal(DocumentAnchorResolutionStatus.NotFound, map.Resolve(block with { ExactText = "伪造文字" }).Status);
    }

    [Fact]
    public void BorrowingAnotherBlockIdCannotOverrideExactSourceIdentity()
    {
        var map = Map(P("目标文字"), P("其他文字"));
        var forgedId = map.Blocks[0] with { BlockId = map.Blocks[1].BlockId };

        var result = map.Resolve(forgedId);

        Assert.True(result.IsResolved);
        Assert.Same(map.Blocks[0], result.Block);
    }

    [Fact]
    public void LargeRepeatedTextMapResolvesThroughUniqueContextWithoutOrdinalGuessing()
    {
        var paragraphs = Enumerable.Range(0, 5000).SelectMany(index => new[] { P("章节上下文 " + index), P("重复正文") }).ToArray();
        var map = Map(paragraphs);

        Assert.Equal(10_000, map.Blocks.Count);
        Assert.Equal(10_000, map.Blocks.Select(block => block.BlockId).Distinct().Count());
        foreach (var block in map.Blocks.Where(block => block.ExactText == "重复正文"))
        {
            var resolved = map.Resolve(block);
            Assert.Equal(DocumentAnchorResolutionStatus.Exact, resolved.Status);
            Assert.Same(block, resolved.Block);
        }
    }

    private static DocumentSourceMap Map(params XElement[] paragraphs)
    {
        var document = new XDocument(new XElement(W + "document", new XElement(W + "body", paragraphs)));
        return DocumentSourceMap.Create("word/document.xml", document.Descendants(W + "p").ToArray());
    }

    private static XElement P(string text, string? paragraphId = null) => new(W + "p",
        paragraphId is null ? null : new XAttribute(W14 + "paraId", paragraphId), Run(text));

    private static XElement Run(string text, string? font = null) => new(W + "r",
        font is null ? null : new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "eastAsia", font))),
        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
}
