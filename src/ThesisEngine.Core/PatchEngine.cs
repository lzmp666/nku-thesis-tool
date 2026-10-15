namespace ThesisEngine.Core;

/// <summary>
/// Stable facade for the plan/apply contract used by CLI and future VSTO code.
/// </summary>
public static class PatchEngine
{
    public static PatchPlan PlanPatches(AnalysisResult result) => PatchPlanner.PlanPatches(result);

    public static ApplyResult ApplyPatches(string inputPath, PatchPlan plan, string outputPath) =>
        WordPatchApplier.ApplyPatches(inputPath, plan, outputPath);

    public static ApplyResult ApplyPatches(string inputPath, IReadOnlyList<Patch> patches, string outputPath) =>
        WordPatchApplier.ApplyPatches(inputPath, patches, outputPath);
}
