using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// A text-bearing direct w:r. Offsets count UTF-16 characters in ExactText, not
/// Word COM character positions; invisible annotation CPs must still be mapped by Word.
/// </summary>
public sealed record DocumentSourceRun(
    int RunIndex,
    int Utf16Offset,
    int Utf16Length,
    string ExactText,
    string TextHash,
    string LegacyLocation);

/// <summary>
/// A snapshot-owned source identity. BlockId survives formatting-only/run-split
/// changes, but is not a permanent identity after repeated paragraphs are inserted.
/// Resolve deliberately does not treat its occurrence ordinal or XML path as proof.
/// </summary>
public sealed record DocumentSourceBlock(
    string BlockId,
    string PartName,
    int ParagraphIndex,
    string? ParagraphId,
    string StructuralPath,
    string ExactText,
    string TextHash,
    string? PreviousTextHash,
    string? NextTextHash,
    IReadOnlyList<DocumentSourceRun> Runs)
{
    public bool HasCompleteTextMapping { get; init; }
    public bool HasUniqueParagraphId { get; init; }
    public int TextOccurrenceCount { get; init; } = 1;
    public int TextOccurrenceIndex { get; init; }
    public IReadOnlyList<string> UnsupportedReasons { get; init; } = Array.Empty<string>();
}

public enum DocumentAnchorResolutionStatus
{
    Exact,
    Relocated,
    Ambiguous,
    TextChanged,
    NotFound
}

public sealed record DocumentAnchorResolution(
    DocumentAnchorResolutionStatus Status,
    DocumentSourceBlock? Block,
    string Reason)
{
    public bool IsResolved => Status is DocumentAnchorResolutionStatus.Exact or DocumentAnchorResolutionStatus.Relocated;
}

public enum DocumentSpanResolutionStatus
{
    Resolved,
    AnchorUnresolved,
    InvalidSpan,
    TextMismatch,
    UnsupportedStructure
}

/// <summary>A verified subspan of one current run, still expressed in XML-text UTF-16 offsets.</summary>
public sealed record DocumentSourceSpanSegment(
    string LegacyLocation,
    int RunIndex,
    int ParagraphUtf16Offset,
    int RunUtf16Offset,
    int Utf16Length,
    string ExactText);

public sealed record DocumentSpanResolution(
    DocumentSpanResolutionStatus Status,
    DocumentSourceBlock? Block,
    IReadOnlyList<DocumentSourceSpanSegment> Segments,
    string Reason)
{
    public bool IsResolved => Status == DocumentSpanResolutionStatus.Resolved;
    public DocumentAnchorResolutionStatus? AnchorStatus { get; init; }
}

/// <summary>
/// Preserves immutable DOCX source locations for structure/AI decisions. Text is
/// never trimmed or normalized. Repeated text requires unchanged, unique context;
/// fuzzy text matching is intentionally unavailable for mutations.
/// </summary>
public sealed class DocumentSourceMap
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace Math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly HashSet<XName> UnsupportedTextElements = new(
        new[] { "tab", "br", "cr", "sym", "noBreakHyphen", "softHyphen", "drawing", "pict", "object",
            "fldSimple", "instrText", "fldChar", "hyperlink", "sdt", "customXml", "altChunk", "smartTag",
            "ins", "del", "moveFrom", "moveTo", "delText", "footnoteReference", "endnoteReference", "txbxContent" }
            .Select(name => W + name).Concat(new[] { Math + "oMath", Math + "oMathPara" }));

    private readonly Dictionary<string, DocumentSourceBlock> _byBlockId;
    private readonly Dictionary<string, DocumentSourceBlock[]> _byParagraphId;
    private readonly Dictionary<string, DocumentSourceBlock[]> _byExactText;
    private readonly Dictionary<TextContext, DocumentSourceBlock[]> _byContext;
    private sealed record TextContext(string ExactText, string? PreviousTextHash, string? NextTextHash);

    private DocumentSourceMap(string partName, IReadOnlyList<DocumentSourceBlock> blocks)
    {
        PartName = partName;
        Blocks = blocks;
        _byBlockId = blocks.ToDictionary(block => block.BlockId, StringComparer.Ordinal);
        _byParagraphId = blocks.Where(block => block.ParagraphId is not null)
            .GroupBy(block => block.ParagraphId!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.OrdinalIgnoreCase);
        _byExactText = blocks.GroupBy(block => block.ExactText, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        _byContext = blocks.GroupBy(block => new TextContext(block.ExactText, block.PreviousTextHash, block.NextTextHash))
            .ToDictionary(group => group.Key, group => group.ToArray());
    }

    public string PartName { get; }
    public IReadOnlyList<DocumentSourceBlock> Blocks { get; }

    /// <summary>
    /// sourceParagraphs must be the reader's document-order w:p list, excluding
    /// text-box stories. No paragraph is silently dropped or reordered.
    /// </summary>
    public static DocumentSourceMap Create(string partName, IReadOnlyList<XElement> sourceParagraphs)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(partName);
        ArgumentNullException.ThrowIfNull(sourceParagraphs);
        if (partName.StartsWith('/') || partName.Contains('\\') || partName.Contains('#') ||
            partName.Split('/').Any(segment => string.IsNullOrWhiteSpace(segment) || segment is "." or ".."))
            throw new ArgumentException("Use a relative OOXML part name, without a fragment or traversal.", nameof(partName));

        var seen = new HashSet<XElement>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in sourceParagraphs)
        {
            if (paragraph is null || paragraph.Name != W + "p" || !seen.Add(paragraph) ||
                paragraph.Ancestors(W + "txbxContent").Any() || paragraph.Ancestors(W + "p").Any())
                throw new ArgumentException("Expected unique main-story w:p elements, excluding nested/text-box paragraphs.", nameof(sourceParagraphs));
        }

        var paths = BuildStructuralPaths(sourceParagraphs);
        var drafts = sourceParagraphs.Select((paragraph, index) => ReadParagraph(partName, index, paragraph, paths[paragraph])).ToArray();
        var textCounts = drafts.GroupBy(draft => draft.ExactText, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var idCounts = drafts.Where(draft => draft.ParagraphId is not null)
            .GroupBy(draft => draft.ParagraphId!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        var previous = new string?[drafts.Length];
        var next = new string?[drafts.Length];
        string? lastHash = null;
        for (var index = 0; index < drafts.Length; index++)
        {
            previous[index] = lastHash;
            if (!string.IsNullOrWhiteSpace(drafts[index].ExactText)) lastHash = drafts[index].TextHash;
        }
        lastHash = null;
        for (var index = drafts.Length - 1; index >= 0; index--)
        {
            next[index] = lastHash;
            if (!string.IsNullOrWhiteSpace(drafts[index].ExactText)) lastHash = drafts[index].TextHash;
        }

        var occurrences = new Dictionary<string, int>(StringComparer.Ordinal);
        var blocks = new DocumentSourceBlock[drafts.Length];
        for (var index = 0; index < drafts.Length; index++)
        {
            var draft = drafts[index];
            var occurrence = occurrences.GetValueOrDefault(draft.ExactText);
            occurrences[draft.ExactText] = occurrence + 1;
            var uniqueId = draft.ParagraphId is not null && idCounts[draft.ParagraphId] == 1;
            var identity = uniqueId ? "paraId:" + draft.ParagraphId : "text:" + draft.TextHash + ":occurrence:" + occurrence.ToString(CultureInfo.InvariantCulture);
            blocks[index] = new DocumentSourceBlock("docx:" + Hash(partName + "\0" + identity), partName, index,
                draft.ParagraphId, draft.StructuralPath, draft.ExactText, draft.TextHash, previous[index], next[index], draft.Runs)
            {
                HasCompleteTextMapping = draft.UnsupportedReasons.Count == 0,
                HasUniqueParagraphId = uniqueId,
                TextOccurrenceCount = textCounts[draft.ExactText],
                TextOccurrenceIndex = occurrence,
                UnsupportedReasons = draft.UnsupportedReasons
            };
        }
        return new DocumentSourceMap(partName, Array.AsReadOnly(blocks));
    }

    public bool TryGetBlock(string blockId, out DocumentSourceBlock? block) => _byBlockId.TryGetValue(blockId, out block);

    /// <summary>
    /// Resolves an old snapshot anchor in this map. A matching stable Word ID with
    /// changed text stops resolution; another copy of that text is not substituted.
    /// </summary>
    public DocumentAnchorResolution Resolve(DocumentSourceBlock anchor)
    {
        ArgumentNullException.ThrowIfNull(anchor);
        if (!string.Equals(anchor.PartName, PartName, StringComparison.Ordinal))
            return Unresolved(DocumentAnchorResolutionStatus.NotFound, "The anchor belongs to a different OOXML part.");
        if (anchor.TextOccurrenceCount < 1 || !string.Equals(anchor.TextHash, Hash(anchor.ExactText), StringComparison.Ordinal))
            return Unresolved(DocumentAnchorResolutionStatus.NotFound, "The source anchor has inconsistent text metadata.");

        if (anchor.HasUniqueParagraphId && anchor.ParagraphId is not null &&
            _byParagraphId.TryGetValue(anchor.ParagraphId, out var byId))
        {
            if (byId.Length != 1)
                return Unresolved(DocumentAnchorResolutionStatus.Ambiguous, "The previously unique Word paragraph ID is now duplicated.");
            if (!SameText(anchor, byId[0]))
                return Unresolved(DocumentAnchorResolutionStatus.TextChanged, "The paragraph with this Word ID has been edited.");
            return Resolved(anchor, byId[0], "Unique Word paragraph ID and exact source text match.");
        }

        if (anchor.ExactText.Length == 0)
            return Unresolved(DocumentAnchorResolutionStatus.Ambiguous, "An empty paragraph requires a unique Word paragraph ID.");
        if (!_byExactText.TryGetValue(anchor.ExactText, out var candidates))
            return Unresolved(DocumentAnchorResolutionStatus.NotFound, "The exact source text is no longer present.");
        if (candidates.Length == 1 && anchor.TextOccurrenceCount == 1)
            return Resolved(anchor, candidates[0], "The exact text is unique in both snapshots.");

        // Inserting/deleting an identical paragraph can make a new occurrence take
        // the old ordinal/path. Neither is a safe identity in that situation.
        if (candidates.Length != anchor.TextOccurrenceCount)
            return Unresolved(DocumentAnchorResolutionStatus.Ambiguous, "The number of identical paragraphs changed; ordinal and XML path cannot identify the original.");

        if (_byContext.TryGetValue(new TextContext(anchor.ExactText, anchor.PreviousTextHash, anchor.NextTextHash), out var contextual) &&
            contextual.Length == 1 && (anchor.PreviousTextHash is not null || anchor.NextTextHash is not null))
            return Resolved(anchor, contextual[0], "Exact text and both nearest nonempty paragraph contexts uniquely match.");
        return Unresolved(DocumentAnchorResolutionStatus.Ambiguous, "Repeated source text has no unique unchanged context.");
    }

    /// <summary>
    /// Reprojects an exact source-text span onto current runs after formatting has
    /// split/merged them. This does not supply or guess Word COM CP coordinates.
    /// </summary>
    public DocumentSpanResolution ResolveSpan(DocumentSourceBlock anchor, int utf16Offset, int utf16Length, string expectedText)
    {
        ArgumentNullException.ThrowIfNull(anchor);
        ArgumentNullException.ThrowIfNull(expectedText);
        if (utf16Offset < 0 || utf16Length <= 0 || utf16Offset > anchor.ExactText.Length - utf16Length ||
            !IsUtf16Boundary(anchor.ExactText, utf16Offset) || !IsUtf16Boundary(anchor.ExactText, utf16Offset + utf16Length))
            return SpanFailure(DocumentSpanResolutionStatus.InvalidSpan, null, "The requested positive UTF-16 span is out of bounds or splits a surrogate pair.");
        if (expectedText.Length != utf16Length ||
            !anchor.ExactText.AsSpan(utf16Offset, utf16Length).SequenceEqual(expectedText.AsSpan()))
            return SpanFailure(DocumentSpanResolutionStatus.TextMismatch, null, "The requested span does not equal the exact source text.");

        var resolution = Resolve(anchor);
        if (!resolution.IsResolved || resolution.Block is null)
            return SpanFailure(DocumentSpanResolutionStatus.AnchorUnresolved, null, resolution.Reason) with { AnchorStatus = resolution.Status };
        var current = resolution.Block;
        if (!anchor.HasCompleteTextMapping || !current.HasCompleteTextMapping)
            return SpanFailure(DocumentSpanResolutionStatus.UnsupportedStructure, current,
                "The paragraph contains nontext/unsupported inline structure; XML text offsets cannot safely represent its Word text.");
        if (!current.ExactText.AsSpan(utf16Offset, utf16Length).SequenceEqual(expectedText.AsSpan()))
            return SpanFailure(DocumentSpanResolutionStatus.TextMismatch, current, "The current paragraph no longer contains the requested exact span.");

        var end = utf16Offset + utf16Length;
        var coveredUntil = utf16Offset;
        var segments = new List<DocumentSourceSpanSegment>();
        foreach (var run in current.Runs)
        {
            var start = System.Math.Max(utf16Offset, run.Utf16Offset);
            var overlapEnd = System.Math.Min(end, run.Utf16Offset + run.Utf16Length);
            if (start >= overlapEnd) continue;
            if (start != coveredUntil)
                return SpanFailure(DocumentSpanResolutionStatus.UnsupportedStructure, current, "The current direct runs do not cover the source span contiguously.");
            segments.Add(new DocumentSourceSpanSegment(run.LegacyLocation, run.RunIndex, start,
                start - run.Utf16Offset, overlapEnd - start, current.ExactText.Substring(start, overlapEnd - start)));
            coveredUntil = overlapEnd;
        }
        if (coveredUntil != end)
            return SpanFailure(DocumentSpanResolutionStatus.UnsupportedStructure, current, "The current direct runs do not cover the complete source span.");
        return new DocumentSpanResolution(DocumentSpanResolutionStatus.Resolved, current, segments.AsReadOnly(),
            "Exact UTF-16 text span verified and projected onto current runs.") { AnchorStatus = resolution.Status };
    }

    private static DocumentAnchorResolution Resolved(DocumentSourceBlock oldBlock, DocumentSourceBlock current, string reason) =>
        new(oldBlock.ParagraphIndex == current.ParagraphIndex && oldBlock.StructuralPath == current.StructuralPath
            ? DocumentAnchorResolutionStatus.Exact : DocumentAnchorResolutionStatus.Relocated, current, reason);

    private static DocumentAnchorResolution Unresolved(DocumentAnchorResolutionStatus status, string reason) => new(status, null, reason);
    private static DocumentSpanResolution SpanFailure(DocumentSpanResolutionStatus status, DocumentSourceBlock? block, string reason) =>
        new(status, block, Array.Empty<DocumentSourceSpanSegment>(), reason);
    private static bool SameText(DocumentSourceBlock first, DocumentSourceBlock second) =>
        string.Equals(first.TextHash, second.TextHash, StringComparison.Ordinal) &&
        string.Equals(first.ExactText, second.ExactText, StringComparison.Ordinal);

    private static bool IsUtf16Boundary(string text, int offset) => offset == 0 || offset == text.Length ||
        !(char.IsHighSurrogate(text[offset - 1]) && char.IsLowSurrogate(text[offset]));

    private static string Hash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();

    private sealed record ParagraphDraft(string? ParagraphId, string StructuralPath, string ExactText, string TextHash,
        IReadOnlyList<DocumentSourceRun> Runs, IReadOnlyList<string> UnsupportedReasons);

    private static ParagraphDraft ReadParagraph(string partName, int paragraphIndex, XElement paragraph, string path)
    {
        var rawId = (string?)paragraph.Attribute(W14 + "paraId");
        var paragraphId = rawId is { Length: 8 } && rawId.All(Uri.IsHexDigit) ? rawId.ToUpperInvariant() : null;
        var textOwners = new Dictionary<XElement, int>(ReferenceEqualityComparer.Instance);
        var textBuilders = new List<StringBuilder>();
        var offsets = new List<int>();
        foreach (var run in paragraph.Elements(W + "r"))
        {
            var texts = run.Descendants(W + "t").ToArray();
            if (texts.Length == 0) continue;
            var runIndex = textBuilders.Count;
            textBuilders.Add(new StringBuilder());
            offsets.Add(-1);
            foreach (var text in texts) textOwners[text] = runIndex;
        }

        var exactText = new StringBuilder();
        var unownedText = false;
        foreach (var text in paragraph.Descendants(W + "t"))
        {
            if (textOwners.TryGetValue(text, out var runIndex))
            {
                if (offsets[runIndex] < 0) offsets[runIndex] = exactText.Length;
                textBuilders[runIndex].Append(text.Value);
            }
            else unownedText = true;
            exactText.Append(text.Value);
        }
        var reasons = paragraph.Descendants().Where(element => UnsupportedTextElements.Contains(element.Name))
            .Select(element => "Unsupported inline element: " + QualifiedName(element.Name)).Distinct(StringComparer.Ordinal).ToList();
        if (unownedText) reasons.Add("Paragraph text contains text outside direct w:r elements.");
        var runs = textBuilders.Select((text, index) =>
        {
            var value = text.ToString();
            return new DocumentSourceRun(index, offsets[index], value.Length, value, Hash(value),
                partName + "#paragraph[" + paragraphIndex.ToString(CultureInfo.InvariantCulture) + "]/run[" + index.ToString(CultureInfo.InvariantCulture) + "]");
        }).ToArray();
        var combined = string.Concat(runs.Select(run => run.ExactText));
        var valueText = exactText.ToString();
        if (!string.Equals(combined, valueText, StringComparison.Ordinal) && !unownedText)
            reasons.Add("Direct runs do not exactly cover paragraph text.");
        if (runs.Any(run => !IsUtf16Boundary(valueText, run.Utf16Offset) || !IsUtf16Boundary(valueText, run.Utf16Offset + run.Utf16Length)))
            reasons.Add("A direct run boundary splits a UTF-16 surrogate pair.");
        return new ParagraphDraft(paragraphId, path, valueText, Hash(valueText), Array.AsReadOnly(runs), reasons.AsReadOnly());
    }

    private static Dictionary<XElement, string> BuildStructuralPaths(IReadOnlyList<XElement> paragraphs)
    {
        var targets = new HashSet<XElement>(paragraphs, ReferenceEqualityComparer.Instance);
        var roots = new HashSet<XElement>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in paragraphs)
            roots.Add(paragraph.Document?.Root ?? paragraph.Ancestors().LastOrDefault() ?? paragraph);
        var result = new Dictionary<XElement, string>(ReferenceEqualityComparer.Instance);
        foreach (var root in roots)
        {
            var pending = new Stack<(XElement Element, string Path)>();
            pending.Push((root, "/" + QualifiedName(root.Name) + "[0]"));
            while (pending.TryPop(out var item))
            {
                if (targets.Contains(item.Element)) result[item.Element] = item.Path;
                // All supplied paragraphs are main-story peers, never nested inside
                // another paragraph. Their run/formatting trees cannot affect paths.
                if (item.Element.Name == W + "p") continue;
                var counts = new Dictionary<XName, int>();
                var children = new List<(XElement Element, string Path)>();
                foreach (var child in item.Element.Elements())
                {
                    var ordinal = counts.GetValueOrDefault(child.Name);
                    counts[child.Name] = ordinal + 1;
                    children.Add((child, item.Path + "/" + QualifiedName(child.Name) + "[" + ordinal.ToString(CultureInfo.InvariantCulture) + "]"));
                }
                for (var index = children.Count - 1; index >= 0; index--) pending.Push(children[index]);
            }
        }
        return result;
    }

    private static string QualifiedName(XName name) => name.Namespace == W ? "w:" + name.LocalName :
        name.Namespace == W14 ? "w14:" + name.LocalName : name.Namespace == Math ? "m:" + name.LocalName :
        name.Namespace == XNamespace.None ? name.LocalName : "{" + name.NamespaceName + "}" + name.LocalName;
}
