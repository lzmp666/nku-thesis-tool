namespace ThesisEngine.Core;

/// <summary>
/// A narrowly scoped OOXML change. Values are the effective values observed by
/// the analyzer, represented as strings so a plan can be reviewed or persisted.
/// </summary>
public sealed record Patch(
    string RuleId,
    string Location,
    string ExpectedBefore,
    string DesiredAfter,
    string Reason);

public sealed record PatchFailure(Patch Patch, string Reason);

/// <summary>
/// The executable part of a repair plan and diagnostics deliberately left for
/// the user. Only auto diagnostics are included in <see cref="Patches"/>.
/// </summary>
public sealed record PatchPlan(
    IReadOnlyList<Patch> Patches,
    IReadOnlyList<Diagnostic> DeferredDiagnostics)
{
    public PatchPlan(IReadOnlyList<Patch> patches)
        : this(patches, Array.Empty<Diagnostic>()) { }

    public static PatchPlan Empty { get; } = new(Array.Empty<Patch>(), Array.Empty<Diagnostic>());
}

public sealed record ApplyResult(
    string? OutputPath,
    IReadOnlyList<Patch> AppliedPatches,
    IReadOnlyList<PatchFailure> Failures,
    IReadOnlyList<Patch> UnresolvedPatches)
{
    public bool Succeeded => Failures.Count == 0;
}
