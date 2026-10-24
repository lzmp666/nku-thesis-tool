using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class WordTableGeometryTests
{
    [Fact]
    public void TwipNormalizationDoesNotBlockObjectUndoButResizingDoes()
    {
        Assert.True(WordTableGeometry.EquivalentDimensions(new[] { 456.1, 114.025 }, new[] { 456.05, 114.05 }));
        Assert.False(WordTableGeometry.EquivalentDimensions(new[] { 456.1, 114.025 }, new[] { 450.0, 114.05 }));
        Assert.False(WordTableGeometry.EquivalentDimensions(new[] { 456.1 }, new[] { 456.1, 1.0 }));
        Assert.False(WordTableGeometry.EquivalentDimensions(new[] { double.NaN }, new[] { double.NaN }));
        Assert.False(WordTableGeometry.EquivalentDimensions(new[] { double.PositiveInfinity }, new[] { double.PositiveInfinity }));
    }

    [Fact]
    public void ConcreteEqualRowsAreAcceptedWithoutDependingOnTheUniformFlag()
    {
        var first = new[] { 100f, 200f, 300f };
        var widths = WordTableGeometry.ReadColumns(new[] { first, new[] { 100f, 200f, 300f } });
        Assert.Equal(first, widths);
        Assert.NotSame(first, widths);
    }

    [Fact]
    public void MergedOrUnequalRowsAreRejected()
    {
        Assert.Throws<InvalidDataException>(() => WordTableGeometry.ReadColumns(new[] { new[] { 100f, 200f }, new[] { 300f } }));
        Assert.Throws<InvalidDataException>(() => WordTableGeometry.ReadColumns(new[] { new[] { 100f, 200f }, new[] { 110f, 190f } }));
    }

    [Theory]
    [InlineData(float.NaN)]
    [InlineData(float.PositiveInfinity)]
    [InlineData(0f)]
    [InlineData(-10f)]
    [InlineData(999999f)]
    public void UndefinedOrInvalidWordWidthsAreRejected(float width)
    {
        Assert.Throws<InvalidDataException>(() => WordTableGeometry.ReadColumns(new[] { new[] { width } }));
    }
}
