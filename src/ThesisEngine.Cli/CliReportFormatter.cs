using System.Globalization;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

public static class CliReportFormatter
{
    public static string Format(AnalysisResult result, string? inputPath = null)
    {
        ArgumentNullException.ThrowIfNull(result);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        Write(result, writer, inputPath);
        return writer.ToString();
    }

    public static void Write(AnalysisResult result, TextWriter writer, string? inputPath = null)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(writer);

        var diagnostics = result.Diagnostics ?? Array.Empty<Diagnostic>();
        var unsupported = result.UnsupportedScopes ?? Array.Empty<string>();

        writer.WriteLine("=== Thesis format analysis ===");
        writer.WriteLine("Mode: DRAFT/DEMO (draft/demo; partial rule coverage)");
        if (!string.IsNullOrWhiteSpace(inputPath))
            writer.WriteLine($"Input: {inputPath}");
        writer.WriteLine("Status: analysis completed");
        writer.WriteLine($"Diagnostics: {diagnostics.Count.ToString(CultureInfo.InvariantCulture)}");
        writer.WriteLine($"Findings: {diagnostics.Count.ToString(CultureInfo.InvariantCulture)}");
        writer.WriteLine($"Rule version: {ValueOrUnknown(result.RuleVersion)}");
        writer.WriteLine($"Coverage: {ValueOrUnknown(result.Coverage)}");

        writer.WriteLine("Unsupported scopes:");
        if (unsupported.Count == 0)
        {
            writer.WriteLine("  (none reported by the service; this demo still covers only a subset of rules)");
        }
        else
        {
            foreach (var scope in unsupported)
                writer.WriteLine($"  - {ValueOrUnknown(scope)}");
        }

        var errorFindings = diagnostics
            .Where(diagnostic => string.Equals(diagnostic.Severity, "error", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (errorFindings.Length > 0)
        {
            writer.WriteLine("Error-severity findings:");
            foreach (var finding in errorFindings)
                writer.WriteLine($"  - {ValueOrUnknown(finding.RuleId)} at {ValueOrUnknown(finding.Location)}: {ValueOrUnknown(finding.Message)}");
        }

        writer.WriteLine("Finding details:");
        if (diagnostics.Count == 0)
        {
            writer.WriteLine("  (none)");
        }
        else
        {
            for (var index = 0; index < diagnostics.Count; index++)
            {
                var diagnostic = diagnostics[index];
                writer.WriteLine($"  [{index + 1}] {ValueOrUnknown(diagnostic.Severity)} {ValueOrUnknown(diagnostic.RuleId)}");
                writer.WriteLine($"      Location: {ValueOrUnknown(diagnostic.Location)}");
                writer.WriteLine($"      Expected: {ValueOrUnknown(diagnostic.Expected)}");
                writer.WriteLine($"      Actual: {ValueOrUnknown(diagnostic.Actual)}");
                writer.WriteLine($"      Fix level: {ValueOrUnknown(diagnostic.FixLevel)}");
                writer.WriteLine($"      Message: {ValueOrUnknown(diagnostic.Message)}");
            }
        }

        writer.WriteLine("Notice: DRAFT/DEMO only; 仅供演示/测试，不代表合规结论。");
    }

    private static string ValueOrUnknown(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "(unknown)" : value;
}
