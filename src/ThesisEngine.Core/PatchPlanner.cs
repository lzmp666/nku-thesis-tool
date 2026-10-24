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

        if (diagnostic.RuleId is FontSizeAnalyzer.RuleIdPictureWidth or FontSizeAnalyzer.RuleIdTableWidth)
        {
            if (!TryParsePoints(diagnostic.Expected, out var limit) || !TryParsePoints(diagnostic.Actual, out var width) || width <= limit)
                return false;
            patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual, diagnostic.Expected,
                "缩小至当前分节可用宽度；图片保持宽高比，简单固定表格保持列宽比例，不修改内容。");
            return true;
        }

        if (string.Equals(diagnostic.RuleId, FontSizeAnalyzer.RuleIdEastAsiaFont, StringComparison.Ordinal))
        {
            patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual,
                diagnostic.Expected, "将正文中的中文字符字体改为目标值；仅供演示规则验证。");
            return true;
        }

        if (string.Equals(diagnostic.RuleId, FontSizeAnalyzer.RuleIdAsciiFont, StringComparison.Ordinal))
        {
            patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual,
                diagnostic.Expected, "将正文中的西文字符字体改为目标值；仅供演示规则验证。");
            return true;
        }

        if (string.Equals(diagnostic.RuleId, FontSizeAnalyzer.RuleIdFontSize, StringComparison.Ordinal) &&
            TryParsePoints(diagnostic.Expected, out _))
        {
            patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual,
                diagnostic.Expected, "将正文中的文字字号改为目标值；仅供演示规则验证。");
            return true;
        }

        if (string.Equals(diagnostic.RuleId, FontSizeAnalyzer.RuleIdCaptionEastAsiaFont, StringComparison.Ordinal) &&
            string.Equals(diagnostic.Expected, "宋体", StringComparison.Ordinal))
        {
            patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual,
                diagnostic.Expected, "将图题或表题中的东亚字体改为宋体；依据 2026 版表 4.6。仍需作者预览确认。");
            return true;
        }

        if (string.Equals(diagnostic.RuleId, FontSizeAnalyzer.RuleIdCaptionFontSize, StringComparison.Ordinal) &&
            TryParsePoints(diagnostic.Expected, out _))
        {
            patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual,
                diagnostic.Expected, "将图题或表题字号改为规范目标；依据 2026 版表 4.6。仍需作者预览确认。");
            return true;
        }

        if (diagnostic.RuleId is FontSizeAnalyzer.RuleIdCaptionAlignment or
            FontSizeAnalyzer.RuleIdCaptionLineSpacing or
            FontSizeAnalyzer.RuleIdCaptionSpaceBefore or
            FontSizeAnalyzer.RuleIdCaptionSpaceAfter)
        {
            var supportedExpected = diagnostic.RuleId switch
            {
                FontSizeAnalyzer.RuleIdCaptionAlignment => string.Equals(diagnostic.Expected, "center", StringComparison.OrdinalIgnoreCase),
                FontSizeAnalyzer.RuleIdCaptionLineSpacing => string.Equals(diagnostic.Expected, "single", StringComparison.OrdinalIgnoreCase),
                _ => TryParseNonNegativePoints(diagnostic.Expected, out _)
            };
            if (supportedExpected)
            {
                patch = new Patch(diagnostic.RuleId, diagnostic.Location, diagnostic.Actual,
                    diagnostic.Expected, "按南开大学 2026 版正文排版要求修改题注段落格式；仅改格式，不移动或重写题名。");
                return true;
            }
        }

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

    internal static bool TryParseNonNegativePoints(string value, out double points)
    {
        points = 0;
        var normalized = value.Trim();
        if (normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            normalized = normalized[..^2].Trim();
        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out points) &&
               points >= 0 && double.IsFinite(points);
    }
}
