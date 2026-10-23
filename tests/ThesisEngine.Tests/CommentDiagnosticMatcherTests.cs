using ThesisAddIn.Protocol;
using ThesisAddIn.VSTO.Ui;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class CommentDiagnosticMatcherTests
{
    [Fact]
    public void ReadsMultilineFieldsWithoutTruncatingGroupedActualValues()
    {
        const string comment = "规则：demo/western-font\r期望值：Times New Roman\r实际值：Cambria；Arial。[NKU-TOOL-ANNOTATION:v1][NKU-LOCATION:p0]";
        Assert.Equal("Times New Roman", CommentDiagnosticMatcher.ReadCommentField(comment, "期望值"));
        Assert.Equal("Cambria；Arial", CommentDiagnosticMatcher.ReadCommentField(comment, "实际值"));
    }

    [Fact]
    public void StillReadsLegacyInlineCommentFields()
    {
        const string comment = "规则：demo/font-size。期望值：12pt；实际值：10.5pt。";
        Assert.Equal("12pt", CommentDiagnosticMatcher.ReadCommentField(comment, "期望值"));
        Assert.Equal("10.5pt", CommentDiagnosticMatcher.ReadCommentField(comment, "实际值"));
    }

    [Fact]
    public void UsesLocationToSelectTheMatchingRepeatedRule()
    {
        var diagnostics = new[]
        {
            Create("word/document.xml#paragraph[0]/run[0]"),
            Create("word/document.xml#paragraph[4]/run[2]")
        };

        Assert.Equal(1, CommentDiagnosticMatcher.FindUniqueMatchIndex(
            diagnostics, "demo/western-font", "Times New Roman", "Arial",
            "word/document.xml#paragraph[4]/run[2]"));
    }

    [Fact]
    public void RejectsLegacyCommentWhenRepeatedDiagnosticsAreAmbiguous()
    {
        var diagnostics = new[]
        {
            Create("word/document.xml#paragraph[0]/run[0]"),
            Create("word/document.xml#paragraph[4]/run[2]")
        };

        Assert.Equal(-1, CommentDiagnosticMatcher.FindUniqueMatchIndex(
            diagnostics, "demo/western-font", "Times New Roman", "Arial", string.Empty));
    }

    [Fact]
    public void DoesNotSelectWhenCommentFieldsDoNotExactlyMatch()
    {
        Assert.Equal(-1, CommentDiagnosticMatcher.FindUniqueMatchIndex(
            new[] { Create("word/document.xml#paragraph[0]/run[0]") },
            "demo/western-font", "times new roman", "Arial",
            "word/document.xml#paragraph[0]/run[0]"));
    }

    [Fact]
    public void MatchesAnyLocationInAGroupedDiagnostic()
    {
        EngineDiagnostic grouped = Create("word/document.xml#paragraph[0]/run[0]");
        grouped.Occurrences = new()
        {
            new EngineDiagnosticOccurrence { Location = grouped.Location, Actual = "Arial" },
            new EngineDiagnosticOccurrence { Location = "word/document.xml#paragraph[3]/run[1]", Actual = "Arial" }
        };

        Assert.Equal(0, CommentDiagnosticMatcher.FindUniqueMatchIndex(
            new[] { grouped }, "demo/western-font", "Times New Roman", "Arial",
            "word/document.xml#paragraph[3]/run[1]"));
    }

    private static EngineDiagnostic Create(string location) => new()
    {
        RuleId = "demo/western-font",
        Expected = "Times New Roman",
        Actual = "Arial",
        Location = location
    };
}

