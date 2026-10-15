using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// Reads the draft/demo font and size snapshot from the main Word document part.
/// The reader is read-only; unresolved values remain null and unsupported parts are reported.
/// </summary>
public sealed class WordFormattingReader : IDemoDocumentReader
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string ThemeNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    public DocumentSnapshot Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!string.Equals(System.IO.Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only .docx files are supported.", nameof(path));

        using var archive = ZipFile.OpenRead(System.IO.Path.GetFullPath(path));
        var document = ReadXml(archive, "word/document.xml", "document");
        var styles = TryReadXml(archive, "word/styles.xml", "styles");
        var theme = TryReadXml(archive, "word/theme/theme1.xml", "theme", ThemeNamespace);
        var defaults = ReadDefaults(styles, theme);
        var styleMap = ReadStyles(styles, theme, defaults.EastAsiaLanguage);
        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidDataException("Missing Word document body.");

        var unsupported = new List<string>();
        if (archive.Entries.Any(e => e.FullName.StartsWith("word/header", StringComparison.Ordinal) || e.FullName.StartsWith("word/footer", StringComparison.Ordinal)) ||
            document.Descendants(W + "headerReference").Any() || document.Descendants(W + "footerReference").Any())
            unsupported.Add("headers/footers");
        if (archive.Entries.Any(e => e.FullName.StartsWith("word/footnotes", StringComparison.Ordinal)))
            unsupported.Add("footnotes");
        if (archive.Entries.Any(e => e.FullName.StartsWith("word/endnotes", StringComparison.Ordinal)))
            unsupported.Add("endnotes");
        if (body.Descendants(W + "tbl").Any()) unsupported.Add("tables");
        if (document.Descendants(W + "txbxContent").Any()) unsupported.Add("textboxes");
        if (body.Descendants(W + "sdt").Any()) unsupported.Add("content controls");
        if (body.Descendants(W + "customXml").Any()) unsupported.Add("custom XML content");
        if (body.Descendants(W + "altChunk").Any()) unsupported.Add("altChunk content");
        if (body.Descendants(W + "hyperlink").Any()) unsupported.Add("hyperlinks");
        if (body.Descendants(W + "smartTag").Any()) unsupported.Add("smart tags");
        if (body.Descendants(W + "fldSimple").Any()) unsupported.Add("simple fields");
        if (new[] { "ins", "del", "moveFrom", "moveTo" }.Any(name => body.Descendants(W + name).Any()))
            unsupported.Add("tracked changes");
        if (body.Descendants(W + "p").Any(p => p.Parent != body && !p.Ancestors(W + "tbl").Any()))
            unsupported.Add("nested body paragraphs");

        var paragraphs = new List<ParagraphSnapshot>();
        var paragraphIndex = 0;
        foreach (var paragraph in body.Elements(W + "p"))
        {
            var paragraphProperties = paragraph.Element(W + "pPr");
            var paragraphStyle = paragraphProperties?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;
            var inherited = Merge(defaults, ResolveStyle(paragraphStyle, styleMap, default, new HashSet<string>(StringComparer.Ordinal), "paragraph"));
            // w:pPr/w:rPr formats the paragraph mark, not the text runs.

            var runs = new List<TextRunSnapshot>();
            var runIndex = 0;
            foreach (var run in paragraph.Elements(W + "r"))
            {
                var properties = run.Element(W + "rPr");
                var characterStyle = properties?.Element(W + "rStyle")?.Attribute(W + "val")?.Value;
                var effective = Merge(inherited, ResolveStyle(characterStyle, styleMap, default, new HashSet<string>(StringComparer.Ordinal), "character"));
                effective = Merge(effective, ReadRunProperties(properties, theme, effective.EastAsiaLanguage));
                var text = string.Concat(run.Descendants(W + "t").Select(t => t.Value));
                runs.Add(new TextRunSnapshot(runIndex++, effective.EastAsia, effective.Ascii, effective.SizePt, text));
            }
            paragraphs.Add(new ParagraphSnapshot(paragraphIndex++, runs));
        }

        if (theme is null && new[] { document, styles }.Where(d => d is not null)
            .SelectMany(d => d!.Descendants(W + "rFonts")).Any(HasThemeFont))
            unsupported.Add("word/theme/theme1.xml");
        else if (new[] { document, styles }.Where(d => d is not null)
                     .SelectMany(d => d!.Descendants(W + "rFonts"))
                     .Any(fonts => HasUnresolvedThemeFont(fonts, theme, defaults.EastAsiaLanguage)))
            unsupported.Add("theme fonts");
        return new DocumentSnapshot("word/document.xml", paragraphs, unsupported.Distinct(StringComparer.Ordinal).ToArray());
    }

    private static XDocument ReadXml(ZipArchive archive, string name, string root)
    {
        var entries = archive.Entries.Where(e => e.FullName == name).ToArray();
        if (entries.Length != 1)
            throw new InvalidDataException(entries.Length == 0 ? $"Missing OOXML part: {name}" : $"Duplicate OOXML part: {name}");
        var entry = entries[0];
        if (entry.Length > 16 * 1024 * 1024) throw new InvalidDataException($"OOXML part too large: {name}");
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings());
        var xml = XDocument.Load(reader);
        if (xml.Root?.Name != W + root) throw new InvalidDataException($"Invalid OOXML root: {name}");
        return xml;
    }

    private static XDocument? TryReadXml(ZipArchive archive, string name, string root, string? namespaceUri = null)
    {
        var entries = archive.Entries.Where(e => e.FullName == name).ToArray();
        if (entries.Length == 0) return null;
        if (entries.Length != 1) throw new InvalidDataException($"Duplicate OOXML part: {name}");
        var entry = entries[0];
        if (entry.Length > 16 * 1024 * 1024) throw new InvalidDataException($"OOXML part too large: {name}");
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings());
        var xml = XDocument.Load(reader);
        if (xml.Root?.Name != XName.Get(root, namespaceUri ?? MainNamespace))
            throw new InvalidDataException($"Invalid OOXML root: {name}");
        return xml;
    }

    private static XmlReaderSettings Settings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        MaxCharactersInDocument = 16 * 1024 * 1024
    };

    private static Formatting ReadDefaults(XDocument? styles, XDocument? theme) =>
        ReadRunProperties(styles?.Root?.Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr"), theme);

    private static Dictionary<string, StyleDefinition> ReadStyles(XDocument? styles, XDocument? theme, string? defaultLanguage)
    {
        var map = new Dictionary<string, StyleDefinition>(StringComparer.Ordinal);
        foreach (var style in styles?.Root?.Elements(W + "style") ?? Enumerable.Empty<XElement>())
        {
            var id = style.Attribute(W + "styleId")?.Value;
            if (string.IsNullOrWhiteSpace(id)) continue;
            map[id] = new StyleDefinition(style.Attribute(W + "type")?.Value,
                style.Element(W + "basedOn")?.Attribute(W + "val")?.Value,
                ReadRunProperties(style.Element(W + "rPr"), theme, defaultLanguage));
        }
        return map;
    }

    private static Formatting ResolveStyle(string? id, IReadOnlyDictionary<string, StyleDefinition> styles, Formatting defaults,
        HashSet<string> seen, string? expectedType)
    {
        if (string.IsNullOrWhiteSpace(id) || !styles.TryGetValue(id, out var style) || !seen.Add(id)) return defaults;
        if (expectedType is not null && style.Type is not null &&
            !string.Equals(style.Type, expectedType, StringComparison.OrdinalIgnoreCase))
            return defaults;
        var parentType = style.Type ?? expectedType;
        return Merge(ResolveStyle(style.BasedOn, styles, defaults, seen, parentType), style.Formatting);
    }

    private static Formatting Merge(Formatting inherited, Formatting direct) => new(
        direct.EastAsiaSpecified ? direct.EastAsia : inherited.EastAsia,
        direct.AsciiSpecified ? direct.Ascii : inherited.Ascii,
        direct.SizeSpecified ? direct.SizePt : inherited.SizePt,
        direct.EastAsiaLanguage ?? inherited.EastAsiaLanguage,
        direct.EastAsiaSpecified || inherited.EastAsiaSpecified,
        direct.AsciiSpecified || inherited.AsciiSpecified,
        direct.SizeSpecified || inherited.SizeSpecified);

    private static Formatting ReadRunProperties(XElement? properties, XDocument? theme, string? fallbackLanguage = null)
    {
        if (properties is null) return default;
        var fonts = properties.Element(W + "rFonts");
        var eastAsiaAttr = fonts?.Attribute(W + "eastAsia")?.Value;
        var eastAsiaTheme = fonts?.Attribute(W + "eastAsiaTheme")?.Value;
        var asciiAttr = fonts?.Attribute(W + "ascii")?.Value ?? fonts?.Attribute(W + "hAnsi")?.Value;
        var asciiTheme = fonts?.Attribute(W + "asciiTheme")?.Value ?? fonts?.Attribute(W + "hAnsiTheme")?.Value;
        var language = properties.Element(W + "lang")?.Attribute(W + "eastAsia")?.Value ?? fallbackLanguage;
        var eastAsia = ResolveFont(eastAsiaAttr, eastAsiaTheme, theme, language);
        var ascii = ResolveFont(asciiAttr, asciiTheme, theme, language);
        var sizeElement = properties.Element(W + "sz");
        var sizeText = sizeElement?.Attribute(W + "val")?.Value;
        double? size = null;
        if (double.TryParse(sizeText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var halfPoints) && halfPoints > 0)
            size = halfPoints / 2.0;
        return new Formatting(eastAsia, ascii, size, language,
            eastAsiaAttr is not null || eastAsiaTheme is not null,
            asciiAttr is not null || asciiTheme is not null,
            sizeElement is not null);
    }

    private static string? ResolveFont(string? direct, string? themeKey, XDocument? theme, string? language)
    {
        if (!string.IsNullOrWhiteSpace(direct)) return direct;
        if (string.IsNullOrWhiteSpace(themeKey) || theme is null) return null;
        var family = themeKey switch
        {
            "majorHAnsi" or "majorEastAsia" => "majorFont",
            "minorHAnsi" or "minorEastAsia" => "minorFont",
            _ => null
        };
        if (family is null) return null;
        var font = theme.Descendants().FirstOrDefault(e => e.Name.LocalName == family);
        var element = themeKey.EndsWith("EastAsia", StringComparison.Ordinal) ? "ea" : "latin";
        var typeface = font?.Elements().FirstOrDefault(e => e.Name.LocalName == element)?.Attribute("typeface")?.Value;
        if (!string.IsNullOrWhiteSpace(typeface)) return typeface;
        if (element != "ea") return null;

        var script = GetThemeScript(language);
        if (script is null) return null;
        return font?.Elements().FirstOrDefault(e => e.Name.LocalName == "font" &&
            string.Equals(e.Attribute("script")?.Value, script, StringComparison.OrdinalIgnoreCase))?.Attribute("typeface")?.Value;
    }

    private static bool HasThemeFont(XElement fonts) => fonts.Attributes().Any(a => a.Name.LocalName.EndsWith("Theme", StringComparison.Ordinal));

    private static bool HasUnresolvedThemeFont(XElement fonts, XDocument? theme, string? fallbackLanguage)
    {
        var language = fonts.Parent?.Element(W + "lang")?.Attribute(W + "eastAsia")?.Value ?? fallbackLanguage;
        return (fonts.Attribute(W + "eastAsiaTheme") is { } eastAsia && ResolveFont(null, eastAsia.Value, theme, language) is null) ||
               (fonts.Attribute(W + "asciiTheme") is { } ascii && ResolveFont(null, ascii.Value, theme, language) is null) ||
               (fonts.Attribute(W + "hAnsiTheme") is { } hAnsi && ResolveFont(null, hAnsi.Value, theme, language) is null);
    }

    private static string? GetThemeScript(string? language)
    {
        if (language is null) return null;
        var normalized = language.Replace('_', '-').ToLowerInvariant();
        if (normalized.StartsWith("zh-tw", StringComparison.Ordinal) || normalized.StartsWith("zh-hk", StringComparison.Ordinal) || normalized.StartsWith("zh-mo", StringComparison.Ordinal))
            return "Hant";
        if (normalized.StartsWith("zh-", StringComparison.Ordinal)) return "Hans";
        return null;
    }

    private readonly record struct Formatting(
        string? EastAsia,
        string? Ascii,
        double? SizePt,
        string? EastAsiaLanguage = null,
        bool EastAsiaSpecified = false,
        bool AsciiSpecified = false,
        bool SizeSpecified = false);
    private sealed record StyleDefinition(string? Type, string? BasedOn, Formatting Formatting);
}
