using System.IO;
using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class WordParagraphValueMapperTests
{
    [Theory]
    [InlineData(0, "left")]
    [InlineData(1, "center")]
    [InlineData(2, "right")]
    [InlineData(3, "both")]
    [InlineData(4, "distribute")]
    public void MapsAlignmentWithoutCollapsingDifferentValues(int value, string expected) =>
        Assert.Equal(expected, WordParagraphValueMapper.Alignment(value));

    [Theory]
    [InlineData(0, 12, "single")]
    [InlineData(1, 18, "multiple:360")]
    [InlineData(2, 24, "multiple:480")]
    [InlineData(3, 14, "atLeast:280")]
    [InlineData(4, 15, "exact:300")]
    [InlineData(5, 18, "multiple:360")]
    [InlineData(5, 12, "single")]
    public void UsesTheReadersOoxmlLineSpacingUnits(int rule, double points, string expected) =>
        Assert.Equal(expected, WordParagraphValueMapper.LineSpacing(rule, points));

    [Fact]
    public void RejectsMixedAndUnsupportedValues()
    {
        Assert.Throws<InvalidDataException>(() => WordParagraphValueMapper.Alignment(9999999));
        Assert.Throws<InvalidDataException>(() => WordParagraphValueMapper.LineSpacing(9999999, 12));
        Assert.Throws<InvalidDataException>(() => WordParagraphValueMapper.LineSpacing(5, 9999999));
    }
}
