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
    private static readonly XNamespace W = WordNamespace;
    private static readonly XNamespace Ct = ContentTypesNamespace;
    private static readonly XNamespace Rel = RelationshipsNamespace;
    private static readonly Regex LocationPattern = new(
        @"\Aword/document\.xml#paragraph\[(?<paragraph>\d+)\]/run\[(?<run>\d+)\]\z",
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

                if (diagnostics.Count == 0)
                {
                    CopyPackage(source, temporaryPath);
                }
                else
                {
                    var targets = ResolveTargets(body, diagnostics);
                    var documentRelationshipsEntry = GetOptionalEntry(entries, "word/_rels/document.xml.rels");
                    var relationships = documentRelationshipsEntry is null
                        ? new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement(Rel + "Relationships"))
                        : ReadXml(documentRelationshipsEntry, Rel + "Relationships", "word/_rels/document.xml.rels");
                    var commentsPartName = ResolveCommentsPartName(relationships);
                    var commentsEntry = GetOptionalEntry(entries, commentsPartName);
                    var comments = commentsEntry is null
                        ? new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement(W + "comments"))
                        : ReadXml(commentsEntry, W + "comments", commentsPartName);
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
                        EnsureHighlighted(target.Run);
                        var starts = new List<XElement>();
                        var ends = new List<XElement>();
                        var references = new List<XElement>();

                        foreach (var diagnostic in target.Diagnostics)
                        {
                            while (usedIds.Contains(nextId)) nextId++;
                            var id = nextId++;
                            usedIds.Add(id);

                            comments.Root.Add(CreateComment(id, timestamp, diagnostic));
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
        foreach (var diagnostic in diagnostics)
        {
            ArgumentNullException.ThrowIfNull(diagnostic);
            if (string.IsNullOrWhiteSpace(diagnostic.RuleId) ||
                diagnostic.Expected is null || diagnostic.Actual is null || diagnostic.Message is null)
                throw new ArgumentException("Diagnostics must include a rule ID, expected value, actual value, and message.", nameof(diagnostics));

            var match = LocationPattern.Match(diagnostic.Location ?? string.Empty);
            if (!match.Success ||
                !int.TryParse(match.Groups["paragraph"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var paragraphIndex) ||
                !int.TryParse(match.Groups["run"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var runIndex))
                throw new NotSupportedException($"Unsupported diagnostic location: {diagnostic.Location}");

            if (paragraphIndex >= paragraphs.Length)
                throw new InvalidDataException($"Diagnostic paragraph does not exist: {diagnostic.Location}");
            var paragraph = paragraphs[paragraphIndex];
            var runs = paragraph.Elements(W + "r").ToArray();
            if (runIndex >= runs.Length)
                throw new InvalidDataException($"Diagnostic run does not exist: {diagnostic.Location}");

            var key = (paragraphIndex, runIndex);
            if (!targetMap.TryGetValue(key, out var target))
            {
                target = new RunTarget(runs[runIndex]);
                targetMap.Add(key, target);
            }
            target.Diagnostics.Add(diagnostic);
        }

        return targetMap.Values.ToArray();
    }

    private static XElement CreateComment(int id, string timestamp, Diagnostic diagnostic)
    {
        var text = $"要求参考来源：2026版第4.7节表4.6；规则：{diagnostic.RuleId}；期望值：{diagnostic.Expected}；实际值：{diagnostic.Actual}。";
        return new XElement(W + "comment",
            new XAttribute(W + "id", id),
            new XAttribute(W + "author", "NKU Thesis Tool"),
            new XAttribute(W + "initials", "NKU"),
            new XAttribute(W + "date", timestamp),
            new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text))));
    }

    private static XElement CreateCommentReference(int id) =>
        new(W + "r",
            new XElement(W + "rPr", new XElement(W + "rStyle", new XAttribute(W + "val", "CommentReference"))),
            new XElement(W + "commentReference", new XAttribute(W + "id", id)));

    private static void EnsureHighlighted(XElement run)
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
            highlight.SetAttributeValue(W + "val", "yellow");
            return;
        }

        highlight = new XElement(W + "highlight", new XAttribute(W + "val", "yellow"));
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

    private sealed class RunTarget(XElement run)
    {
        public XElement Run { get; } = run;
        public List<Diagnostic> Diagnostics { get; } = [];
    }
}
