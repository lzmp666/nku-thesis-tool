using ThesisEngine.Core.Analysis;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class PageNumberEstimatorTests
{
    private static string Fixture(string relative)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return Path.Combine(directory.FullName, "tests", "fixtures", "synthetic", relative);
    }

    [Fact]
    public void FullThesisGetsMonotonicEstimatedPages()
    {
        var map = new PageNumberEstimator().Estimate(Fixture("full-thesis/synthetic_complete_thesis.docx"));

        Assert.NotNull(map.TotalPages);
        Assert.True(map.TotalPages > 1);
        Assert.NotEmpty(map.Paragraphs);
        Assert.All(map.Paragraphs, p => Assert.NotNull(p.StartPage));
        for (var i = 1; i < map.Paragraphs.Count; i++)
            Assert.True(map.Paragraphs[i - 1].StartPage <= map.Paragraphs[i].StartPage);
        Assert.All(map.Paragraphs, p => Assert.True(p.EndPage >= p.StartPage));
    }

    [Fact]
    public void EstimateContainsForcedBreaksAndIsHonestAboutAccuracy()
    {
        var map = new PageNumberEstimator().Estimate(Fixture("full-thesis/synthetic_complete_thesis.docx"));

        Assert.Contains(map.Paragraphs, p => p.ForcedBreakBefore);
        Assert.Contains("估算", map.Note);
        Assert.Contains("未经 Word 实机渲染", map.Note, StringComparison.Ordinal);
    }

    [Fact]
    public void CompliantSampleGetsValidPageMap()
    {
        var map = new PageNumberEstimator().Estimate(Fixture("analysis/analysis_compliant.docx"));

        Assert.NotEmpty(map.Paragraphs);
        Assert.All(map.Paragraphs, p =>
        {
            if (p.StartPage is int page) Assert.True(page >= 1);
        });
    }
}
