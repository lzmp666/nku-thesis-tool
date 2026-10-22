using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// Writes diagnostic highlights and comments to a new copy of a DOCX package.
/// </summary>
public static class WordAnnotationWriter
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string CommentsRelationshipType = OfficeRelationshipsNamespace + "/comments";
    private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    private const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";
    private const long MaximumXmlCharacters = 16 * 1024 * 1024;
    private const string ToolAnnotationMarker = "[NKU-TOOL-ANNOTATION:v1]";
    private static readonly XNamespace W = WordNamespace;
    private static readonly XNamespace Ct = ContentTypesNamespace;
    private static readonly XNamespace Rel = RelationshipsNamespace;
    private static readonly Regex LocationPattern = new(
        @"\Aword/document\.xml#paragraph\[(?<paragraph>\d+)\]/run\[(?<run>\d+)\]\z",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex LocationMarkerPattern = new(
        @"\[NKU-LOCATION:(?<location>word/document\.xml#paragraph\[\d+\]/run\[\d+\])\]",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex OccurrenceMarkerPattern = new(
        @"\[NKU-OCCURRENCE:(?<location>word/document\.xml#paragraph\[\d+\]/run\[\d+\])\|(?<highlight>[a-zA-Z]+|absent)\]",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// Creates a copy of <paramref name="inputPath"/> with diagnostic highlights and comments.
    /// Existing output files are never overwritten.
    /// </summary>
    public static string WriteAnnotatedCopy(string inputPath, string outputPath, IReadOnlyList<Diagnostic> diagnostics)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(diagnostics);

        EnsureDocxPath(inputPath, nameof(inputPath));
        EnsureDocxPath(outputPath, nameof(outputPath));
        var fullInputPath = Path.GetFullPath(inputPath);
        var fullOutputPath = Path.GetFullPath(outputPath);
        var pathComparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        if (string.Equals(fullInputPath, fullOutputPath, pathComparison))
            throw new ArgumentException("The output path must differ from the input path.", nameof(outputPath));
        if (!File.Exists(fullInputPath)) throw new FileNotFoundException("The input document was not found.", fullInputPath);
        if (File.Exists(fullOutputPath)) throw new IOException("The output file already exists.");

        var outputDirectory = Path.GetDirectoryName(fullOutputPath)!;
        var temporaryPath = Path.Combine(outputDirectory, $".{Path.GetFileNameWithoutExtension(fullOutputPath)}.{Guid.NewGuid():N}.tmp.docx");

        try
        {
            using (var source = ZipFile.OpenRead(fullInputPath))
            {
                EnsureUniqueEntryNames(source);
                var entries = source.Entries.ToDictionary(entry => entry.FullName, StringComparer.Ordinal);
                var documentEntry = GetRequiredEntry(entries, "word/document.xml");
                var contentTypesEntry = GetRequiredEntry(entries, "[Content_Types].xml");
                GetRequiredEntry(entries, "_rels/.rels");

                var document = ReadXml(documentEntry, W + "document", "word/document.xml");
                var body = document.Root!.Element(W + "body")
                    ?? throw new InvalidDataException("Missing Word document body.");

                var targets = ResolveTargets(body, diagnostics);
                var targetByLocation = targets.ToDictionary(target => target.Location, StringComparer.Ordinal);
                var documentRelationshipsEntry = GetOptionalEntry(entries, "word/_rels/document.xml.rels");
                var relationships = documentRelationshipsEntry is null
                    ? new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement(Rel + "Relationships"))
                    : ReadXml(documentRelationshipsEntry, Rel + "Relationships", "word/_rels/document.xml.rels");
                var commentsPartName = ResolveCommentsPartName(relationships);
                var commentsEntry = GetOptionalEntry(entries, commentsPartName);
                var comments = commentsEntry is null
                    ? new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement(W + "comments"))
                    : ReadXml(commentsEntry, W + "comments", commentsPartName);
                var ownedAnnotations = ReadOwnedAnnotations(comments);

                if (targets.Count == 0 && ownedAnnotations.Count == 0)
                {
                    CopyPackage(source, temporaryPath);
                }
                else
                {
                    // Remove only comments carrying our private ownership marker. User comments,
                    // including comments authored as "NKU Thesis Tool", remain untouched.
                    RemoveOwnedAnnotations(document, body, comments, ownedAnnotations);
                    var contentTypes = ReadXml(contentTypesEntry, Ct + "Types", "[Content_Types].xml");
                    EnsureCommentsRelationship(relationships, commentsPartName);
                    EnsureRelationshipsContentType(contentTypes);
                    EnsureCommentsContentType(contentTypes, commentsPartName);

                    var usedIds = comments.Root!.Elements(W + "comment")
                        .Select(comment => comment.Attribute(W + "id")?.Value)
                        .Select(value => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ? id : (int?)null)
                        .Where(id => id.HasValue)
                        .Select(id => id!.Value)
                        .ToHashSet();
                    var nextId = 0;
                    var timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

                    foreach (var target in targets)
                    {
                        var originalHighlight = GetHighlightValue(target.Run);
                        EnsureHighlighted(target.Run);
                        var starts = new List<XElement>();
                        var ends = new List<XElement>();
                        var references = new List<XElement>();

                        foreach (var diagnostic in target.Diagnostics)
                        {
                            while (usedIds.Contains(nextId)) nextId++;
                            var id = nextId++;
                            usedIds.Add(id);

                            var diagnosticOccurrences = GetOccurrences(diagnostic);
                            var occurrenceHighlights = diagnosticOccurrences.Select(occurrence =>
                            {
                                if (!targetByLocation.TryGetValue(occurrence.Location, out var occurrenceTarget))
                                    throw new InvalidDataException("A grouped diagnostic occurrence was not resolved: " + occurrence.Location);
                                return new AnnotationHighlight(occurrence.Location, occurrenceTarget.OriginalHighlight);
                            }).ToArray();
                            comments.Root.Add(CreateComment(id, timestamp, diagnostic, originalHighlight, occurrenceHighlights));
                            starts.Add(new XElement(W + "commentRangeStart", new XAttribute(W + "id", id)));
                            ends.Insert(0, new XElement(W + "commentRangeEnd", new XAttribute(W + "id", id)));
                            references.Add(CreateCommentReference(id));
                        }

                        target.Run.AddBeforeSelf(starts);
                        target.Run.AddAfterSelf(ends.Concat(references));
                    }

                    var modifiedParts = new Dictionary<string, XDocument>(StringComparer.Ordinal)
                    {
                        ["word/document.xml"] = document,
                        ["word/_rels/document.xml.rels"] = relationships,
                        ["[Content_Types].xml"] = contentTypes,
                        [commentsPartName] = comments
                    };
                    WritePackage(source, temporaryPath, modifiedParts);
                }
            }

            DocumentLoader.Load(temporaryPath);
            File.Move(temporaryPath, fullOutputPath);
            return fullOutputPath;
        }
        finally
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    private static void EnsureDocxPath(string path, string parameterName)
    {
        if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only .docx files are supported.", parameterName);
    }

    private static IReadOnlyList<RunTarget> ResolveTargets(XElement body, IReadOnlyList<Diagnostic> diagnostics)
    {
        var paragraphs = body.Elements(W + "p").ToArray();
        var targetMap = new Dictionary<(int Paragraph, int Run), RunTarget>();
        var targetByLocation = new Dictionary<string, RunTarget>(StringComparer.Ordinal);
        foreach (var diagnostic in diagnostics)
        {
            ArgumentNullException.ThrowIfNull(diagnostic);
            if (string.IsNullOrWhiteSpace(diagnostic.RuleId) ||
                diagnostic.Expected is null || diagnostic.Actual is null || diagnostic.Message is null)
                throw new ArgumentException("Diagnostics must include a rule ID, expected value, actual value, and message.", nameof(diagnostics));

            foreach (var occurrence in GetOccurrences(diagnostic))
            {
                var (paragraphIndex, runIndex) = ParseRunLocation(occurrence.Location);
                if (paragraphIndex >= paragraphs.Length)
                    throw new InvalidDataException($"Diagnostic paragraph does not exist: {occurrence.Location}");
                var paragraph = paragraphs[paragraphIndex];
                // Locations count text-bearing runs only. Comment-reference runs are structural
                // OOXML and must not shift a diagnostic when the document is analyzed again.
                var runs = paragraph.Elements(W + "r")
                    .Where(run => run.Descendants(W + "t").Any())
                    .ToArray();
                if (runIndex >= runs.Length)
                    throw new InvalidDataException($"Diagnostic run does not exist: {occurrence.Location}");

                var key = (paragraphIndex, runIndex);
                if (!targetMap.TryGetValue(key, out var target))
                {
                    target = new RunTarget(runs[runIndex], occurrence.Location, GetHighlightValue(runs[runIndex]) ?? "absent");
                    targetMap.Add(key, target);
                }
                targetByLocation[occurrence.Location] = target;
                if (string.Equals(occurrence.Location, diagnostic.Location, StringComparison.Ordinal))
                    target.Diagnostics.Add(diagnostic);
            }
        }

        return targetMap.Values.ToArray();
    }

    private static XElement CreateComment(int id, string timestamp, Diagnostic diagnostic, string? originalHighlight,
        IReadOnlyList<AnnotationHighlight> occurrenceHighlights)
    {
        var text = $"要求参考来源：{GetSourceReference(diagnostic)}；规则：{DescribeRule(diagnostic.RuleId)}；期望值：{diagnostic.Expected}；实际值：{diagnostic.Actual}。";
        return new XElement(W + "comment",
            new XAttribute(W + "id", id),
            new XAttribute(W + "author", "NKU Thesis Tool"),
            new XAttribute(W + "initials", "NKU"),
            new XAttribute(W + "date", timestamp),
            new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)),
                // Hidden metadata keeps the visible Word comment to four concise, user-facing fields.
                CreateHiddenMetadata(ToolAnnotationMarker),
                CreateHiddenMetadata("[NKU-LOCATION:" + diagnostic.Location + "]"),
                CreateHiddenMetadata("[NKU-HIGHLIGHT:" + (originalHighlight ?? "absent") + "]"),
                occurrenceHighlights.Select(occurrence => CreateHiddenMetadata(
                    "[NKU-OCCURRENCE:" + occurrence.Location + "|" + occurrence.Highlight + "]"))));
    }

    private static string GetSourceReference(Diagnostic diagnostic)
    {
        if (diagnostic.RuleId.StartsWith("NKU-REQ-", StringComparison.Ordinal))
        {
            var separator = diagnostic.Message?.IndexOf('：') ?? -1;
            if (separator > 0) return diagnostic.Message![..separator];
        }

        return "演示规则（仅用于验证插件流程，不代表正式标准）";
    }

    private static string DescribeRule(string ruleId) => ruleId switch
    {
        "demo/east-asian-font" or "NKU-REQ-042" => "中文字体",
        "demo/western-font" or "NKU-REQ-043" => "西文字体",
        "demo/font-size" or "NKU-REQ-044" => "字号",
        _ => "格式项目"
    };

    private static XElement CreateHiddenMetadata(string value) =>
        new(W + "r",
            new XElement(W + "rPr", new XElement(W + "vanish")),
            new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value));

    private static IReadOnlyList<OwnedAnnotation> ReadOwnedAnnotations(XDocument comments)
    {
        var owned = new List<OwnedAnnotation>();
        foreach (var comment in comments.Root!.Elements(W + "comment"))
        {
            var hiddenText = string.Concat(comment.Descendants(W + "r")
                .Where(run => run.Element(W + "rPr")?.Element(W + "vanish") is not null)
                .SelectMany(run => run.Elements(W + "t").Select(text => text.Value)));
            if (comment.Attribute(W + "author")?.Value != "NKU Thesis Tool" ||
                !hiddenText.Contains(ToolAnnotationMarker, StringComparison.Ordinal)) continue;

            var id = comment.Attribute(W + "id")?.Value;
            var locationMatch = LocationMarkerPattern.Match(hiddenText);
            var location = locationMatch.Success ? locationMatch.Groups["location"].Value : null;
            var highlight = ReadMetadata(hiddenText, "[NKU-HIGHLIGHT:");
            if (!int.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out _) || location is null || highlight is null ||
                (!string.Equals(highlight, "absent", StringComparison.Ordinal) &&
                 !Regex.IsMatch(highlight, @"\A[a-zA-Z]+\z", RegexOptions.CultureInvariant)))
                throw new InvalidDataException("An NKU Thesis Tool comment contains invalid ownership metadata.");

            var occurrences = OccurrenceMarkerPattern.Matches(hiddenText)
                .Cast<Match>()
                .Select(match => new AnnotationHighlight(match.Groups["location"].Value, match.Groups["highlight"].Value))
                .ToArray();
            if (occurrences.Length == 0)
                occurrences = new[] { new AnnotationHighlight(location, highlight) };
            owned.Add(new OwnedAnnotation(id, occurrences));
        }
        return owned;
    }

    private static string? ReadMetadata(string text, string prefix)
    {
        var start = text.IndexOf(prefix, StringComparison.Ordinal);
        if (start < 0) return null;
        start += prefix.Length;
        var end = text.IndexOf(']', start);
        return end < 0 ? null : text[start..end];
    }

    private static void RemoveOwnedAnnotations(
        XDocument document,
        XElement body,
        XDocument comments,
        IReadOnlyList<OwnedAnnotation> ownedAnnotations)
    {
        if (ownedAnnotations.Count == 0) return;
        var highlightByLocation = ownedAnnotations
            .SelectMany(annotation => annotation.Occurrences)
            .GroupBy(occurrence => occurrence.Location, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group =>
            {
                var values = group.Select(annotation => annotation.Highlight).Distinct(StringComparer.Ordinal).ToArray();
                if (values.Length != 1)
                    throw new InvalidDataException("Owned comments disagree about the original highlight at " + group.Key + ".");
                return values[0];
            }, StringComparer.Ordinal);
        var ids = ownedAnnotations.Select(annotation => annotation.Id).ToHashSet(StringComparer.Ordinal);

        foreach (var location in highlightByLocation)
        {
            var run = ResolveTextRun(body, location.Key);
            RestoreHighlight(run, location.Value);
        }

        foreach (var marker in document.Descendants().Where(element =>
                     (element.Name == W + "commentRangeStart" || element.Name == W + "commentRangeEnd") &&
                     ids.Contains(element.Attribute(W + "id")?.Value ?? string.Empty)).ToArray())
            marker.Remove();

        foreach (var reference in document.Descendants(W + "commentReference")
                     .Where(element => ids.Contains(element.Attribute(W + "id")?.Value ?? string.Empty)).ToArray())
        {
            var run = reference.Ancestors(W + "r").FirstOrDefault();
            if (run is not null && run.Elements().All(child => child.Name == W + "rPr" || child == reference))
                run.Remove();
            else
                reference.Remove();
        }

        foreach (var comment in comments.Root!.Elements(W + "comment")
                     .Where(element => ids.Contains(element.Attribute(W + "id")?.Value ?? string.Empty)).ToArray())
            comment.Remove();
    }

    private static XElement ResolveTextRun(XElement body, string location)
    {
        var match = LocationPattern.Match(location);
        if (!match.Success ||
            !int.TryParse(match.Groups["paragraph"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var paragraphIndex) ||
            !int.TryParse(match.Groups["run"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var runIndex))
            throw new InvalidDataException("An owned comment has an unsupported location: " + location);
        var paragraphs = body.Elements(W + "p").ToArray();
        if (paragraphIndex >= paragraphs.Length)
            throw new InvalidDataException("An owned comment paragraph no longer exists: " + location);
        var runs = paragraphs[paragraphIndex].Elements(W + "r")
            .Where(run => run.Descendants(W + "t").Any())
            .ToArray();
        if (runIndex >= runs.Length)
            throw new InvalidDataException("An owned comment text run no longer exists: " + location);
        return runs[runIndex];
    }

    private static IReadOnlyList<DiagnosticOccurrence> GetOccurrences(Diagnostic diagnostic) =>
        diagnostic.Occurrences.Count > 0
            ? diagnostic.Occurrences
            : new[] { new DiagnosticOccurrence(diagnostic.Location, diagnostic.Actual) };

    private static (int Paragraph, int Run) ParseRunLocation(string location)
    {
        var match = LocationPattern.Match(location ?? string.Empty);
        if (!match.Success ||
            !int.TryParse(match.Groups["paragraph"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var paragraphIndex) ||
            !int.TryParse(match.Groups["run"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var runIndex))
            throw new NotSupportedException($"Unsupported diagnostic location: {location}");
        return (paragraphIndex, runIndex);
    }

    private static string? GetHighlightValue(XElement run) =>
        run.Element(W + "rPr")?.Element(W + "highlight")?.Attribute(W + "val")?.Value;

    private static void RestoreHighlight(XElement run, string value)
    {
        var highlight = run.Element(W + "rPr")?.Element(W + "highlight");
        if (string.Equals(value, "absent", StringComparison.Ordinal))
        {
            highlight?.Remove();
            return;
        }
        SetHighlightValue(run, value);
    }

    private static XElement CreateCommentReference(int id) =>
        new(W + "r",
            new XElement(W + "rPr", new XElement(W + "rStyle", new XAttribute(W + "val", "CommentReference"))),
            new XElement(W + "commentReference", new XAttribute(W + "id", id)));

    private static void EnsureHighlighted(XElement run)
        => SetHighlightValue(run, "yellow");

    private static void SetHighlightValue(XElement run, string value)
    {
        var properties = run.Element(W + "rPr");
        if (properties is null)
        {
            properties = new XElement(W + "rPr");
            run.AddFirst(properties);
        }

        var highlight = properties.Element(W + "highlight");
        if (highlight is not null)
        {
            highlight.SetAttributeValue(W + "val", value);
            return;
        }

        highlight = new XElement(W + "highlight", new XAttribute(W + "val", value));
        var followingProperty = properties.Elements().FirstOrDefault(element => RunPropertyOrder(element.Name) > 250);
        if (followingProperty is not null)
            followingProperty.AddBeforeSelf(highlight);
        else if (properties.Element(W + "rPrChange") is { } change)
            change.AddBeforeSelf(highlight);
        else
            properties.Add(highlight);
    }

    private static int RunPropertyOrder(XName name) => name == W + "rStyle" ? 0
        : name == W + "rFonts" ? 10
        : name == W + "b" ? 20
        : name == W + "bCs" ? 30
        : name == W + "i" ? 40
        : name == W + "iCs" ? 50
        : name == W + "caps" ? 60
        : name == W + "smallCaps" ? 70
        : name == W + "strike" ? 80
        : name == W + "dstrike" ? 90
        : name == W + "outline" ? 100
        : name == W + "shadow" ? 110
        : name == W + "emboss" ? 120
        : name == W + "imprint" ? 130
        : name == W + "noProof" ? 140
        : name == W + "snapToGrid" ? 150
        : name == W + "vanish" ? 160
        : name == W + "webHidden" ? 170
        : name == W + "color" ? 180
        : name == W + "spacing" ? 190
        : name == W + "w" ? 200
        : name == W + "kern" ? 210
        : name == W + "position" ? 220
        : name == W + "sz" ? 230
        : name == W + "szCs" ? 240
        : name == W + "highlight" ? 250
        : name == W + "u" ? 260
        : name == W + "effect" ? 270
        : name == W + "bdr" ? 280
        : name == W + "shd" ? 290
        : name == W + "fitText" ? 300
        : name == W + "vertAlign" ? 310
        : name == W + "rtl" ? 320
        : name == W + "cs" ? 330
        : name == W + "em" ? 340
        : name == W + "lang" ? 350
        : name == W + "eastAsianLayout" ? 360
        : name == W + "specVanish" ? 370
        : name == W + "oMath" ? 380
        : name == W + "rPrChange" ? 1000
        : -1;

    private static string ResolveCommentsPartName(XDocument relationships)
    {
        var root = relationships.Root!;
        var allIds = root.Elements(Rel + "Relationship")
            .Select(element => element.Attribute("Id")?.Value)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToArray();
        if (allIds.Distinct(StringComparer.Ordinal).Count() != allIds.Length)
            throw new InvalidDataException("Duplicate relationship IDs in word/_rels/document.xml.rels.");

        var commentRelationships = root.Elements(Rel + "Relationship")
            .Where(element => element.Attribute("Type")?.Value == CommentsRelationshipType)
            .ToArray();
        if (commentRelationships.Length > 1)
            throw new InvalidDataException("Multiple comments relationships in word/_rels/document.xml.rels.");
        if (commentRelationships.Length == 1)
        {
            var relationship = commentRelationships[0];
            if (relationship.Attribute("TargetMode")?.Value == "External")
                throw new InvalidDataException("The comments relationship cannot target an external resource.");
            var target = relationship.Attribute("Target")?.Value
                ?? throw new InvalidDataException("The comments relationship has no target.");
            return ResolvePartTarget(target);
        }

        return "word/comments.xml";
    }

    private static string ResolvePartTarget(string target)
    {
        if (!Uri.TryCreate(new Uri("http://package/word/document.xml"), target, out var resolved) ||
            resolved.Host != "package" || resolved.Query.Length > 0 || resolved.Fragment.Length > 0)
            throw new InvalidDataException("Invalid comments relationship target.");

        var partName = Uri.UnescapeDataString(resolved.AbsolutePath.TrimStart('/'));
        if (string.IsNullOrWhiteSpace(partName) || partName.Contains('\\') || partName.Split('/').Any(segment => segment is "" or "." or ".."))
            throw new InvalidDataException("Invalid comments relationship target.");
        return partName;
    }

    private static void EnsureCommentsRelationship(XDocument relationships, string commentsPartName)
    {
        var root = relationships.Root!;
        var existing = root.Elements(Rel + "Relationship")
            .Where(element => element.Attribute("Type")?.Value == CommentsRelationshipType)
            .SingleOrDefault();
        if (existing is not null) return;

        var usedIds = root.Elements(Rel + "Relationship")
            .Select(element => element.Attribute("Id")?.Value)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToHashSet(StringComparer.Ordinal);
        var idNumber = 1;
        while (usedIds.Contains($"rId{idNumber}")) idNumber++;
        var target = new Uri("http://package/word/document.xml")
            .MakeRelativeUri(new Uri("http://package/" + commentsPartName))
            .ToString();
        root.Add(new XElement(Rel + "Relationship",
            new XAttribute("Id", $"rId{idNumber}"),
            new XAttribute("Type", CommentsRelationshipType),
            new XAttribute("Target", target)));
    }

    private static void EnsureRelationshipsContentType(XDocument contentTypes)
    {
        var defaults = contentTypes.Root!.Elements(Ct + "Default")
            .Where(element => string.Equals(element.Attribute("Extension")?.Value, "rels", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (defaults.Length > 1) throw new InvalidDataException("Duplicate rels content type defaults.");
        if (defaults.Length == 1)
        {
            if (defaults[0].Attribute("ContentType")?.Value != RelationshipsContentType)
                throw new InvalidDataException("Invalid rels content type default.");
            return;
        }

        var defaultElement = new XElement(Ct + "Default",
            new XAttribute("Extension", "rels"),
            new XAttribute("ContentType", RelationshipsContentType));
        if (contentTypes.Root.Elements(Ct + "Override").FirstOrDefault() is { } firstOverride)
            firstOverride.AddBeforeSelf(defaultElement);
        else
            contentTypes.Root.Add(defaultElement);
    }

    private static void EnsureCommentsContentType(XDocument contentTypes, string commentsPartName)
    {
        var partName = "/" + commentsPartName;
        var overrides = contentTypes.Root!.Elements(Ct + "Override")
            .Where(element => string.Equals(element.Attribute("PartName")?.Value, partName, StringComparison.Ordinal))
            .ToArray();
        if (overrides.Length > 1) throw new InvalidDataException($"Duplicate content type override: {partName}");
        if (overrides.Length == 1)
        {
            if (overrides[0].Attribute("ContentType")?.Value != CommentsContentType)
                throw new InvalidDataException($"Invalid comments content type: {partName}");
            return;
        }

        contentTypes.Root.Add(new XElement(Ct + "Override",
            new XAttribute("PartName", partName),
            new XAttribute("ContentType", CommentsContentType)));
    }

    private static void WritePackage(ZipArchive source, string outputPath, IReadOnlyDictionary<string, XDocument> modifiedParts)
    {
        using var output = ZipFile.Open(outputPath, ZipArchiveMode.Create);
        var written = new HashSet<string>(StringComparer.Ordinal);
        foreach (var sourceEntry in source.Entries)
        {
            if (modifiedParts.TryGetValue(sourceEntry.FullName, out var modified))
            {
                WriteXml(output, sourceEntry.FullName, modified, sourceEntry);
                written.Add(sourceEntry.FullName);
            }
            else
            {
                CopyEntry(sourceEntry, output);
            }
        }

        foreach (var (partName, xml) in modifiedParts)
        {
            if (!written.Contains(partName)) WriteXml(output, partName, xml, null);
        }
    }

    private static void CopyPackage(ZipArchive source, string outputPath)
    {
        using var output = ZipFile.Open(outputPath, ZipArchiveMode.Create);
        foreach (var entry in source.Entries) CopyEntry(entry, output);
    }

    private static void CopyEntry(ZipArchiveEntry sourceEntry, ZipArchive output)
    {
        var target = output.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
        target.LastWriteTime = sourceEntry.LastWriteTime;
        target.ExternalAttributes = sourceEntry.ExternalAttributes;
        using var source = sourceEntry.Open();
        using var destination = target.Open();
        source.CopyTo(destination);
    }

    private static void WriteXml(ZipArchive output, string partName, XDocument document, ZipArchiveEntry? sourceEntry)
    {
        var entry = output.CreateEntry(partName, CompressionLevel.Optimal);
        if (sourceEntry is not null)
        {
            entry.LastWriteTime = sourceEntry.LastWriteTime;
            entry.ExternalAttributes = sourceEntry.ExternalAttributes;
        }

        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration is null
        });
        document.Save(writer);
    }

    private static XDocument ReadXml(ZipArchiveEntry entry, XName expectedRoot, string partName)
    {
        if (entry.Length > MaximumXmlCharacters) throw new InvalidDataException($"OOXML part too large: {partName}");
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaximumXmlCharacters
        });
        var xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        if (xml.Root?.Name != expectedRoot) throw new InvalidDataException($"Invalid OOXML root: {partName}");
        return xml;
    }

    private static void EnsureUniqueEntryNames(ZipArchive archive)
    {
        var duplicate = archive.Entries.GroupBy(entry => entry.FullName, StringComparer.Ordinal)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) throw new InvalidDataException($"Duplicate OOXML part: {duplicate.Key}");
    }

    private static ZipArchiveEntry GetRequiredEntry(IReadOnlyDictionary<string, ZipArchiveEntry> entries, string partName) =>
        entries.TryGetValue(partName, out var entry)
            ? entry
            : throw new InvalidDataException($"Missing OOXML part: {partName}");

    private static ZipArchiveEntry? GetOptionalEntry(IReadOnlyDictionary<string, ZipArchiveEntry> entries, string partName) =>
        entries.TryGetValue(partName, out var entry) ? entry : null;

    private sealed class RunTarget(XElement run, string location, string originalHighlight)
    {
        public XElement Run { get; } = run;
        public string Location { get; } = location;
        public string OriginalHighlight { get; } = originalHighlight;
        public List<Diagnostic> Diagnostics { get; } = [];
    }

    private sealed record AnnotationHighlight(string Location, string Highlight);
    private sealed record OwnedAnnotation(string Id, IReadOnlyList<AnnotationHighlight> Occurrences);
}

