using ThesisAddIn.Protocol;
using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DiagnosticSetComparerTests
{
    private static EngineDiagnostic Item(string location = "p1", string actual = "Arial") => new() {
        RuleId = "font", Expected = "Times New Roman", Actual = actual, Location = location,
        FixLevel = "auto", Severity = "warning", Message = "source",
        Occurrences = new() { new() { Location = location, Actual = actual } }
    };
    [Fact]
    public void UnchangedDiagnosticsCanSkipAnnotationRebuildRegardlessOfListOrder()
    {
        Assert.True(DiagnosticSetComparer.Equivalent(new[] { Item("a"), Item("b") }, new[] { Item("b"), Item("a") }));
    }
    [Fact]
    public void AnyNewOccurrenceOrChangedValueOrRuleBlocksFastPath()
    {
        Assert.False(DiagnosticSetComparer.Equivalent(new[] { Item() }, new[] { Item("p2") }));
        Assert.False(DiagnosticSetComparer.Equivalent(new[] { Item() }, new[] { Item(actual:"Cambria") }));
        var other = Item(); other.FixLevel = "manual";
        Assert.False(DiagnosticSetComparer.Equivalent(new[] { Item() }, new[] { other }));
        other = Item(); other.Occurrences.Add(new() { Location="p2", Actual="Arial" });
        Assert.False(DiagnosticSetComparer.Equivalent(new[] { Item() }, new[] { other }));
        Assert.False(DiagnosticSetComparer.Equivalent(new[] { Item() }, Array.Empty<EngineDiagnostic>()));
    }
}
