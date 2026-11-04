namespace ThesisEngine.Core;

/// <summary>
/// Collapses run-level findings with the same rule and target value into one issue,
/// while retaining every run and its actual value for navigation and repair.
/// </summary>
public static class DiagnosticGrouper
{
    public static IReadOnlyList<Diagnostic> Group(IEnumerable<Diagnostic> diagnostics)
    {
        ArgumentNullException.ThrowIfNull(diagnostics);

        return diagnostics
            .GroupBy(diagnostic => new
            {
                diagnostic.RuleId,
                diagnostic.Severity,
                diagnostic.Expected,
                diagnostic.FixLevel
            })
            .Select(group =>
            {
                var items = group.ToArray();
                var occurrences = items.SelectMany(diagnostic =>
                        diagnostic.Occurrences.Count > 0
                            ? diagnostic.Occurrences
                            : new[] { new DiagnosticOccurrence(diagnostic.Location, diagnostic.Actual) })
                    .GroupBy(occurrence => occurrence.Location, StringComparer.Ordinal)
                    .Select(locationGroup => locationGroup.First())
                    .ToArray();
                var actualValues = occurrences.Select(occurrence => occurrence.Actual)
                    .Distinct(StringComparer.Ordinal)
                    .ToArray();
                var actual = string.Join("、", actualValues);
                var first = items[0];
                var suffix = occurrences.Length > 1 ? $"（共 {occurrences.Length} 处，统一要求：{first.Expected}。）" : string.Empty;

                return first with
                {
                    Location = occurrences.Length > 0 ? occurrences[0].Location : first.Location,
                    Actual = actual,
                    Message = first.Message + suffix,
                    Occurrences = occurrences
                };
            })
            .ToArray();
    }
}
