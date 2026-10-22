using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// Applies a complete patch batch to a new DOCX copy. All preconditions are
/// checked before any output is published, so a conflict cannot create a
/// partially repaired document.
/// </summary>
public static class WordPatchApplier
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const long MaximumXmlCharacters = 16 * 1024 * 1024;
    private static readonly XNamespace W = WordNamespace;

    public static ApplyResult ApplyPatches(string inputPath, PatchPlan plan, string outputPath) =>
        ApplyPatches(inputPath, plan?.Patches ?? throw new ArgumentNullException(nameof(plan)), outputPath);

    public static ApplyResult ApplyPatches(string inputPath, IReadOnlyList<Patch> patches, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(patches);
        EnsureDocxPath(inputPath, nameof(inputPath));
        EnsureDocxPath(outputPath, nameof(outputPath));

        var fullInputPath = Path.GetFullPath(inputPath);
        var fullOutputPath = Path.GetFullPath(outputPath);
        var comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        if (string.Equals(fullInputPath, fullOutputPath, comparison))
            throw new ArgumentException("The output path must differ from the input path.", nameof(outputPath));
        if (!File.Exists(fullInputPath)) throw new FileNotFoundException("The input document was not found.", fullInputPath);
        if (File.Exists(fullOutputPath)) throw new IOException("The output file already exists.");

        var outputDirectory = Path.GetDirectoryName(fullOutputPath)!;
        Directory.CreateDirectory(outputDirectory);
        var temporaryPath = Path.Combine(outputDirectory, $".{Path.GetFileNameWithoutExtension(fullOutputPath)}.{Guid.NewGuid():N}.tmp.docx");
        var applied = Array.Empty<Patch>();

        try
        {
            using var source = ZipFile.OpenRead(fullInputPath);
            EnsureUniqueEntryNames(source);
            var entries = source.Entries.ToDictionary(entry => entry.FullName, StringComparer.Ordinal);
            var documentEntry = GetRequiredEntry(entries, "word/document.xml");
            GetRequiredEntry(entries, "[Content_Types].xml");
            GetRequiredEntry(entries, "_rels/.rels");
            var document = ReadXml(documentEntry, W + "document", "word/document.xml");
            var body = document.Root?.Element(W + "body")
                ?? throw new InvalidDataException("Missing Word document body.");

            var snapshot = patches.Count == 0 ? null : new WordFormattingReader().Read(fullInputPath);
            var targets = ResolveTargets(body, patches, snapshot);
            ValidateAll(patches, targets, snapshot);
            foreach (var target in targets)
                ApplyTarget(target);

            WritePackage(source, temporaryPath, patches.Count == 0
                ? null
                : new Dictionary<string, XDocument>(StringComparer.Ordinal) { ["word/document.xml"] = document });

            // Re-open the generated package before publication. This catches a
            // malformed XML/package write while the destination is still temp.
            DocumentLoader.Load(temporaryPath);
            File.Move(temporaryPath, fullOutputPath);
            applied = patches.ToArray();
            return new ApplyResult(fullOutputPath, applied, Array.Empty<PatchFailure>(), Array.Empty<Patch>());
        }
        catch (PatchFailureException exception)
        {
            return new ApplyResult(null, Array.Empty<Patch>(), [new PatchFailure(exception.Patch, exception.Message)], patches);
        }
        catch (InvalidDataException exception) when (patches.Count > 0)
        {
            var failed = patches[0];
            return new ApplyResult(null, Array.Empty<Patch>(), [new PatchFailure(failed, exception.Message)], patches);
        }
        catch (NotSupportedException exception) when (patches.Count > 0)
        {
            var failed = patches[0];
            return new ApplyResult(null, Array.Empty<Patch>(), [new PatchFailure(failed, exception.Message)], patches);
        }
        catch
        {
            // The caller gets the original exception and no destination is
            // published. The finally block removes the temporary package.
            throw;
        }
        finally
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    public static ApplyResult Apply(string inputPath, PatchPlan plan, string outputPath) =>
        ApplyPatches(inputPath, plan, outputPath);

    private static void ValidateAll(IReadOnlyList<Patch> patches, IReadOnlyList<PatchTarget> targets, DocumentSnapshot? snapshot)
    {
        var seen = new HashSet<(string Rule, string Location)>();
        foreach (var (patch, target) in patches.Zip(targets))
        {
            if (!seen.Add((patch.RuleId, patch.Location)))
                throw new PatchFailureException(patch, $"Duplicate patch: {patch.RuleId} at {patch.Location}");
            if (target.Snapshot is null || snapshot is null)
                throw new PatchFailureException(patch, $"Unable to resolve patch location: {patch.Location}");
            var actual = target.Snapshot;
            if (string.Equals(patch.RuleId, FontSizeAnalyzer.RuleIdEastAsiaFont, StringComparison.Ordinal))
            {
                EnsureExpected(patch, actual.EastAsiaFont, patch.ExpectedBefore);
            }
            else if (string.Equals(patch.RuleId, FontSizeAnalyzer.RuleIdAsciiFont, StringComparison.Ordinal))
            {
                EnsureExpected(patch, DescribeWesternFonts(actual), patch.ExpectedBefore);
            }
            else if (string.Equals(patch.RuleId, FontSizeAnalyzer.RuleIdFontSize, StringComparison.Ordinal))
            {
                EnsureExpected(patch, actual.FontSizePt is null ? null : $"{actual.FontSizePt.Value.ToString(CultureInfo.InvariantCulture)}pt", patch.ExpectedBefore);
                if (!PatchPlanner.TryParsePoints(patch.DesiredAfter, out var points) || !TryHalfPoints(points, out _))
                    throw new PatchFailureException(patch, $"Invalid desired font size: {patch.DesiredAfter}");
            }
            else
                throw new PatchFailureException(patch, $"Unsupported patch rule: {patch.RuleId}");
        }
    }

    private static void EnsureExpected(Patch patch, string? actual, string expected)
    {
        if (actual is null || string.Equals(actual, "(unknown)", StringComparison.OrdinalIgnoreCase) ||
            !ValuesEqual(actual, expected, patch.RuleId))
            throw new PatchFailureException(patch, $"Patch precondition failed at {patch.Location}: expected {expected}, actual {actual ?? "(unknown)"}.");
    }

    private static bool ValuesEqual(string actual, string expected, string ruleId)
    {
        if (string.Equals(ruleId, FontSizeAnalyzer.RuleIdFontSize, StringComparison.Ordinal))
            return PatchPlanner.TryParsePoints(actual, out var actualPoints) &&
                   PatchPlanner.TryParsePoints(expected, out var expectedPoints) &&
                   Math.Abs(actualPoints - expectedPoints) <= 0.001;
        return string.Equals(actual, expected, StringComparison.Ordinal);
    }

    private static IReadOnlyList<PatchTarget> ResolveTargets(XElement body, IReadOnlyList<Patch> patches, DocumentSnapshot? snapshot)
    {
        var paragraphs = body.Descendants(W + "p")
            .Where(paragraph => !paragraph.Ancestors(W + "txbxContent").Any())
            .ToArray();
        var result = new List<PatchTarget>(patches.Count);
        foreach (var patch in patches)
        {
            ArgumentNullException.ThrowIfNull(patch);
            var (paragraphIndex, runIndex) = ParseLocation(patch);
            if (paragraphIndex >= paragraphs.Length)
                throw new PatchFailureException(patch, $"Patch paragraph does not exist: {patch.Location}");
            // Locations count text-bearing runs, just as the reader and annotation
            // writer do. Comment-reference runs must never become patch targets.
            var runs = paragraphs[paragraphIndex].Elements(W + "r")
                .Where(run => run.Descendants(W + "t").Any()).ToArray();
            if (runIndex >= runs.Length)
                throw new PatchFailureException(patch, $"Patch run does not exist: {patch.Location}");
            var effective = snapshot?.Paragraphs.FirstOrDefault(p => p.ParagraphIndex == paragraphIndex)?.Runs
                .FirstOrDefault(r => r.RunIndex == runIndex);
            result.Add(new PatchTarget(patch, runs[runIndex]) { Snapshot = effective });
        }
        return result;
    }

    private static (int Paragraph, int Run) ParseLocation(Patch patch)
    {
        const string prefix = "word/document.xml#paragraph[";
        var location = patch.Location;
        if (!location.StartsWith(prefix, StringComparison.Ordinal) ||
            !location.Contains("]/run[", StringComparison.Ordinal) || !location.EndsWith(']'))
            throw new PatchFailureException(patch, $"Unsupported patch location: {location}");
        var separator = location.IndexOf("]/run[", StringComparison.Ordinal);
        var pText = location[prefix.Length..separator];
        var rText = location[(separator + "]/run[".Length)..^1];
        if (!int.TryParse(pText, NumberStyles.None, CultureInfo.InvariantCulture, out var paragraph) || paragraph < 0 ||
            !int.TryParse(rText, NumberStyles.None, CultureInfo.InvariantCulture, out var run) || run < 0)
            throw new PatchFailureException(patch, $"Unsupported patch location: {location}");
        return (paragraph, run);
    }

    private static void ApplyTarget(PatchTarget target)
    {
        var properties = target.Run.Element(W + "rPr");
        if (properties is null)
        {
            properties = new XElement(W + "rPr");
            target.Run.AddFirst(properties);
        }

        if (string.Equals(target.Patch.RuleId, FontSizeAnalyzer.RuleIdEastAsiaFont, StringComparison.Ordinal))
        {
            var fonts = GetOrCreateRunFonts(properties);
            fonts.SetAttributeValue(W + "eastAsia", target.Patch.DesiredAfter);
            fonts.Attribute(W + "eastAsiaTheme")?.Remove();
        }
        else if (string.Equals(target.Patch.RuleId, FontSizeAnalyzer.RuleIdAsciiFont, StringComparison.Ordinal))
        {
            var fonts = GetOrCreateRunFonts(properties);
            fonts.SetAttributeValue(W + "ascii", target.Patch.DesiredAfter);
            fonts.SetAttributeValue(W + "hAnsi", target.Patch.DesiredAfter);
            // Explicit values must win over a theme slot after the repair.
            fonts.Attribute(W + "asciiTheme")?.Remove();
            fonts.Attribute(W + "hAnsiTheme")?.Remove();
        }
        else if (string.Equals(target.Patch.RuleId, FontSizeAnalyzer.RuleIdFontSize, StringComparison.Ordinal))
        {
            if (!PatchPlanner.TryParsePoints(target.Patch.DesiredAfter, out var points) || !TryHalfPoints(points, out var halfPoints))
                throw new PatchFailureException(target.Patch, $"Invalid desired font size: {target.Patch.DesiredAfter}");
            var size = properties.Element(W + "sz");
            if (size is null)
            {
                size = new XElement(W + "sz");
                var followingProperty = properties.Elements().FirstOrDefault(element => IsAfterFontSize(element.Name));
                if (followingProperty is null) properties.Add(size);
                else followingProperty.AddBeforeSelf(size);
            }
            size.SetAttributeValue(W + "val", halfPoints.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static XElement GetOrCreateRunFonts(XElement properties)
    {
        var fonts = properties.Element(W + "rFonts");
        if (fonts is not null) return fonts;

        fonts = new XElement(W + "rFonts");
        var style = properties.Element(W + "rStyle");
        if (style is null) properties.AddFirst(fonts);
        else style.AddAfterSelf(fonts);
        return fonts;
    }

    private static string DescribeWesternFonts(TextRunSnapshot snapshot)
    {
        var ascii = snapshot.AsciiSlotFont ?? snapshot.AsciiFont;
        var hAnsi = snapshot.HAnsiFont ?? snapshot.AsciiFont;
        var normalizedAscii = string.IsNullOrWhiteSpace(ascii) ? "(unknown)" : ascii;
        var normalizedHAnsi = string.IsNullOrWhiteSpace(hAnsi) ? "(unknown)" : hAnsi;
        return string.Equals(normalizedAscii, normalizedHAnsi, StringComparison.Ordinal)
            ? normalizedAscii
            : $"ASCII={normalizedAscii}; HAnsi={normalizedHAnsi}";
    }

    private static bool IsAfterFontSize(XName name) =>
        name == W + "szCs" || name == W + "highlight" || name == W + "u" || name == W + "effect" ||
        name == W + "bdr" || name == W + "shd" || name == W + "fitText" || name == W + "vertAlign" ||
        name == W + "rtl" || name == W + "cs" || name == W + "em" || name == W + "lang" ||
        name == W + "eastAsianLayout" || name == W + "specVanish" || name == W + "oMath" || name == W + "rPrChange";

    private static bool TryHalfPoints(double points, out int halfPoints)
    {
        halfPoints = (int)Math.Round(points * 2, MidpointRounding.AwayFromZero);
        return points > 0 && Math.Abs(points * 2 - halfPoints) < 0.0001 && halfPoints > 0;
    }

    private static void WritePackage(ZipArchive source, string outputPath, IReadOnlyDictionary<string, XDocument>? modifiedParts)
    {
        using var output = ZipFile.Open(outputPath, ZipArchiveMode.Create);
        foreach (var entry in source.Entries)
        {
            if (modifiedParts is not null && modifiedParts.TryGetValue(entry.FullName, out var document))
                WriteXml(output, entry.FullName, document, entry);
            else
                CopyEntry(entry, output);
        }
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

    private static void WriteXml(ZipArchive output, string partName, XDocument document, ZipArchiveEntry sourceEntry)
    {
        var entry = output.CreateEntry(partName, CompressionLevel.Optimal);
        entry.LastWriteTime = sourceEntry.LastWriteTime;
        entry.ExternalAttributes = sourceEntry.ExternalAttributes;
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(false),
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
        var duplicate = archive.Entries.GroupBy(e => e.FullName, StringComparer.Ordinal).FirstOrDefault(g => g.Count() > 1);
        if (duplicate is not null) throw new InvalidDataException($"Duplicate OOXML part: {duplicate.Key}");
    }

    private static ZipArchiveEntry GetRequiredEntry(IReadOnlyDictionary<string, ZipArchiveEntry> entries, string partName) =>
        entries.TryGetValue(partName, out var entry) ? entry : throw new InvalidDataException($"Missing OOXML part: {partName}");

    private static void EnsureDocxPath(string path, string parameterName)
    {
        if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only .docx files are supported.", parameterName);
    }

    private sealed class PatchTarget(Patch patch, XElement run)
    {
        public Patch Patch { get; } = patch;
        public XElement Run { get; } = run;
        public TextRunSnapshot? Snapshot { get; set; }
    }

    private sealed class PatchFailureException(Patch patch, string message) : Exception(message)
    {
        public Patch Patch { get; } = patch;
    }
}
