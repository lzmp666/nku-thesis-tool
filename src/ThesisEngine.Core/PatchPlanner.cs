using System.Globalization;

namespace ThesisEngine.Core;

/// <summary>
/// Converts diagnostics into executable, reviewable patches. Suggestions and
/// manual findings remain in DeferredDiagnostics and are never executable.
/// </summary>
public static class PatchPlanner
{
    public static PatchPlan PlanPatches(AnalysisResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var patches = new List<Patch>();
        var deferred = new List<Diagnostic>();

        foreach (var diagnostic in result.Diagnostics ?? Array.Empty<Diagnostic>())
        {
            var occurrences = diagnostic.Occurrences.Count > 0
                ? diagnostic.Occurrences
                : new[] { new DiagnosticOccurrence(diagnostic.Location, diagnostic.Actual) };

            if (!string.Equals(diagnostic.FixLevel, "auto", StringComparison.OrdinalIgnoreCase))
            {
                deferred.Add(diagnostic);
                continue;
            }

            var created = new List<Patch>();
            var allValid = true;
            foreach (var occurrence in occurrences)
            {
                var occurrenceDiagnostic = diagnostic with
                {
                    Location = occurrence.Location,
                    Actual = occurrence.Actual,
                    Occurrences = Array.Empty<DiagnosticOccurrence>()
                };
                if (TryCreatePatch(occurrenceDiagnostic, out var patch))
                    created.Add(patch!);
                else
                    allValid = false;
            }

            if (allValid && created.Count == occurrences.Count)
                patches.AddRange(created);
            else
                deferred.Add(diagnostic);
        }

        return new PatchPlan(patches, deferred);
    }

    public static PatchPlan Plan(AnalysisResult result) => PlanPatches(result);

    private static bool TryCreatePatch(Diagnostic diagnostic, out Patch? patch)
    {
        patch = null;
        if (string.IsNullOrWhiteSpace(diagnostic.RuleId) ||
            string.IsNullOrWhiteSpace(diagnostic.Location) ||
            string.IsNullOrWhiteSpace(diagnostic.Expected) ||
            string.IsNullOrWhiteSpace(diagnostic.Actual) ||
            string.Equals(diagnostic.Actual, "(unknown)", StringComparison.OrdinalIgnoreCase))
            return false;

        if (string.Equals(diagnostic.RuleId, FontSizeAnalyzer.RuleIdEastAsiaFont, StringComparison.Ordinal))
        {
            patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual,
                diagnostic.Expected, "将正文 run 的东亚字体改为规则目标值。");
            return true;
        }

        if (string.Equals(diagnostic.RuleId, FontSizeAnalyzer.RuleIdFontSize, StringComparison.Ordinal) &&
            TryParsePoints(diagnostic.Expected, out _))
        {
            patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual,
                diagnostic.Expected, "将正文 run 的字号改为规则目标值。");
            return true;
        }

        // ASCII font remains suggest/manual until its safety and scope are
        // explicitly approved; auto plans must not silently broaden scope.
        return false;
    }

    internal static bool TryParsePoints(string value, out double points)
    {
        points = 0;
        var normalized = value.Trim();
        if (normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            normalized = normalized[..^2].Trim();
        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out points) &&
               points > 0 && double.IsFinite(points);
    }
}
