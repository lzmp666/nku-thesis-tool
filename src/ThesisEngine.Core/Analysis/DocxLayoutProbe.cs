using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core.Analysis;

public sealed record SectionLayout(
    int Index,
    string PageSize,
    double? WidthCm,
    double? HeightCm,
    double? TopCm,
    double? BottomCm,
    double? LeftCm,
    double? RightCm,
    double? HeaderCm,
    double? FooterCm,
    double? GutterCm,
    int FirstParagraphIndex,
    int LastParagraphIndex);

public sealed record RunLayout(
    int RunIndex,
    string Text,
    bool? Bold,
    bool BoldKnown);

public sealed record ParagraphLayout(
    int ParagraphIndex,
    int SectionIndex,
    string Alignment,
    double SpaceBeforePt,
    double SpaceAfterPt,
    string LineSpacingKind,
    double? LineSpacingPt,
    double? LineSpacingMultiple,
    double? FirstLineChars,
    double? FirstLinePt,
    double? HangingChars,
    double? HangingPt,
    bool HasEquation,
    bool ValuesKnown,
    IReadOnlyList<RunLayout> Runs);

public sealed record DocumentLayoutProbe(
    IReadOnlyList<SectionLayout> Sections,
    IReadOnlyList<ParagraphLayout> Paragraphs)
{
    public ParagraphLayout? TryGetParagraph(int index) =>
        index >= 0 && index < Paragraphs.Count ? Paragraphs[index] : null;
}

/// <summary>
/// Reads the paragraph properties that the demo reader does not expose: first-line
/// indent, section page setup, and effective bold. It does not duplicate font or
/// size resolution; those continue to come from <see cref="WordFormattingReader"/>.
/// </summary>
public sealed class DocxLayoutProbe
{
    private const double TwipsPerCentimetre = 566.9291338582677;
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace MathNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public DocumentLayoutProbe Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(Path.GetFullPath(path), FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var document = ReadXml(archive, "word/document.xml");
        var styles = TryReadXml(archive, "word/styles.xml");
        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidDataException("Missing Word document body.");

        var styleMap = ReadStyles(styles);
        var defaultParagraphStyle = styles?.Root?.Elements(W + "style")
            .FirstOrDefault(style => (string?)style.Attribute(W + "type") == "paragraph" &&
                new[] { "1", "true", "on" }.Contains((string?)style.Attribute(W + "default"), StringComparer.OrdinalIgnoreCase))
            ?.Attribute(W + "styleId")?.Value;
        var defaultsPPr = ReadPPr(styles?.Root?.Element(W + "docDefaults")?
            .Element(W + "pPrDefault")?.Element(W + "pPr"));
        var defaultsBold = ReadBold(styles?.Root?.Element(W + "docDefaults")?
            .Element(W + "rPrDefault")?.Element(W + "rPr"));

        var paragraphs = body.Descendants(W + "p")
            .Where(paragraph => !paragraph.Ancestors(W + "txbxContent").Any())
            .ToArray();
        var layouts = new List<ParagraphLayout>(paragraphs.Length);
        var breaks = new List<(int EndParagraph, XElement? SectPr)>();
        for (var index = 0; index < paragraphs.Length; index++)
        {
            var paragraph = paragraphs[index];
            var properties = paragraph.Element(W + "pPr");
            var styleId = properties?.Element(W + "pStyle")?.Attribute(W + "val")?.Value
                ?? defaultParagraphStyle;
            var styleResolution = ResolvePPr(styleId, styleMap, defaultsPPr, new HashSet<string>(StringComparer.Ordinal));
            var paragraphStyleBold = ResolveStyleBold(styleId, styleMap, defaultsBold, new HashSet<string>(StringComparer.Ordinal));
            var effective = Merge(styleResolution.Values, ReadPPr(properties));
            var textRuns = paragraph.Elements(W + "r").Where(run => run.Descendants(W + "t").Any()).ToArray();
            var runs = new List<RunLayout>(textRuns.Length);
            for (var runIndex = 0; runIndex < textRuns.Length; runIndex++)
            {
                var run = textRuns[runIndex];
                var runProperties = run.Element(W + "rPr");
                var characterStyle = runProperties?.Element(W + "rStyle")?.Attribute(W + "val")?.Value;
                var characterBold = ResolveStyleBold(characterStyle, styleMap, null, new HashSet<string>(StringComparer.Ordinal));
                var directBold = ReadBold(runProperties);
                var bold = directBold ?? characterBold.Values ?? paragraphStyleBold.Values;
                var boldKnown = bold is not null && characterBold.IsResolved && styleResolution.IsResolved;
                runs.Add(new RunLayout(runIndex,
                    string.Concat(run.Descendants(W + "t").Select(t => t.Value)),
                    bold, boldKnown));
            }

            layouts.Add(new ParagraphLayout(
                index,
                -1,
                effective.Alignment ?? "left",
                effective.BeforePt ?? 0,
                effective.AfterPt ?? 0,
                effective.Line?.Kind ?? "single",
                effective.Line?.Pt,
                effective.Line?.Multiple,
                effective.FirstLineChars,
                effective.FirstLinePt,
                effective.HangingChars,
                effective.HangingPt,
                paragraph.Descendants(MathNs + "oMath").Any() || paragraph.Descendants(MathNs + "oMathPara").Any(),
                styleResolution.IsResolved,
                runs.AsReadOnly()));

            var sectionProperties = properties?.Element(W + "sectPr");
            if (sectionProperties is not null)
                breaks.Add((index, sectionProperties));
        }

        var bodySection = body.Element(W + "sectPr");
        if (bodySection is not null && (breaks.Count == 0 || breaks[^1].EndParagraph != paragraphs.Length - 1))
            breaks.Add((paragraphs.Length - 1, bodySection));

        var sections = BuildSections(breaks, paragraphs.Length);
        var sectionByParagraph = new int[paragraphs.Length];
        foreach (var section in sections)
        {
            for (var index = section.FirstParagraphIndex; index <= section.LastParagraphIndex && index < sectionByParagraph.Length; index++)
                sectionByParagraph[index] = section.Index;
        }

        var withSections = layouts.Select(layout => layout with { SectionIndex = sectionByParagraph[layout.ParagraphIndex] }).ToArray();
        return new DocumentLayoutProbe(sections.ToArray(), Array.AsReadOnly<ParagraphLayout>(withSections));
    }

    private static IReadOnlyList<SectionLayout> BuildSections(List<(int EndParagraph, XElement? SectPr)> breaks, int paragraphCount)
    {
        var sections = new List<SectionLayout>();
        if (paragraphCount == 0)
        {
            sections.Add(ReadSection(0, 0, -1, null));
            return sections.AsReadOnly();
        }

        breaks.Sort((left, right) => left.EndParagraph.CompareTo(right.EndParagraph));
        var start = 0;
        var index = 0;
        foreach (var item in breaks)
        {
            if (item.EndParagraph < start) continue;
            sections.Add(ReadSection(index++, start, item.EndParagraph, item.SectPr));
            start = item.EndParagraph + 1;
        }
        if (start < paragraphCount)
            sections.Add(ReadSection(index, start, paragraphCount - 1, null));
        if (sections.Count == 0)
            sections.Add(ReadSection(0, 0, paragraphCount - 1, null));
        return sections.AsReadOnly();
    }

    private static SectionLayout ReadSection(int index, int firstParagraph, int lastParagraph, XElement? sectPr)
    {
        var size = sectPr?.Element(W + "pgSz");
        var width = ParseTwips(size?.Attribute(W + "w")?.Value) / 20;
        var height = ParseTwips(size?.Attribute(W + "h")?.Value) / 20;
        var widthCm = width is null ? null : width / 72.0 * 2.54;
        var heightCm = height is null ? null : height / 72.0 * 2.54;
        var pageSize = widthCm is not null && heightCm is not null && IsA4(widthCm.Value, heightCm.Value) ? "A4" : "闈濧4";
        var margins = sectPr?.Element(W + "pgMar");
        return new SectionLayout(
            index,
            pageSize,
            widthCm,
            heightCm,
            Centimetres(margins?.Attribute(W + "top")?.Value),
            Centimetres(margins?.Attribute(W + "bottom")?.Value),
            Centimetres(margins?.Attribute(W + "left")?.Value),
            Centimetres(margins?.Attribute(W + "right")?.Value),
            Centimetres(margins?.Attribute(W + "header")?.Value),
            Centimetres(margins?.Attribute(W + "footer")?.Value),
            Centimetres(margins?.Attribute(W + "gutter")?.Value),
            firstParagraph,
            lastParagraph);
    }

    private static bool IsA4(double widthCm, double heightCm) =>
        Math.Abs(widthCm - 21.0) <= 0.1 && Math.Abs(heightCm - 29.7) <= 0.1 ||
        Math.Abs(widthCm - 29.7) <= 0.1 && Math.Abs(heightCm - 21.0) <= 0.1;

    private static double? Centimetres(string? twips) => ParseTwips(twips) is { } value ? value / TwipsPerCentimetre : null;

    private static double? ParseTwips(string? value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed : null;

    private static Dictionary<string, StyleDefinition> ReadStyles(XDocument? styles)
    {
        var map = new Dictionary<string, StyleDefinition>(StringComparer.Ordinal);
        foreach (var style in styles?.Root?.Elements(W + "style") ?? Enumerable.Empty<XElement>())
        {
            var id = style.Attribute(W + "styleId")?.Value;
            if (string.IsNullOrWhiteSpace(id)) continue;
            map[id] = new StyleDefinition(
                style.Attribute(W + "type")?.Value,
                style.Element(W + "basedOn")?.Attribute(W + "val")?.Value,
                ReadPPr(style.Element(W + "pPr")),
                ReadBold(style.Element(W + "rPr")));
        }
        return map;
    }

    private static ParagraphResolution ResolvePPr(
        string? id,
        IReadOnlyDictionary<string, StyleDefinition> styles,
        PPr defaults,
        HashSet<string> seen)
    {
        if (string.IsNullOrWhiteSpace(id)) return new ParagraphResolution(defaults, true);
        if (!styles.TryGetValue(id, out var style) || !seen.Add(id))
            return new ParagraphResolution(defaults, false);
        var parent = ResolvePPr(style.BasedOn, styles, defaults, seen);
        return new ParagraphResolution(Merge(parent.Values, style.Paragraph), parent.IsResolved);
    }

    private static BoldResolution ResolveStyleBold(
        string? id,
        IReadOnlyDictionary<string, StyleDefinition> styles,
        bool? defaults,
        HashSet<string> seen)
    {
        if (string.IsNullOrWhiteSpace(id)) return new BoldResolution(defaults, true);
        if (!styles.TryGetValue(id, out var style) || !seen.Add(id))
            return new BoldResolution(defaults, false);
        var parent = ResolveStyleBold(style.BasedOn, styles, defaults, seen);
        return new BoldResolution(style.Bold ?? parent.Values, parent.IsResolved);
    }

    private static PPr Merge(PPr inherited, PPr direct) => new(
        direct.Alignment ?? inherited.Alignment,
        direct.BeforePt ?? inherited.BeforePt,
        direct.AfterPt ?? inherited.AfterPt,
        direct.Line ?? inherited.Line,
        direct.FirstLineChars ?? inherited.FirstLineChars,
        direct.FirstLinePt ?? inherited.FirstLinePt,
        direct.HangingChars ?? inherited.HangingChars,
        direct.HangingPt ?? inherited.HangingPt);

    private static PPr ReadPPr(XElement? properties)
    {
        if (properties is null) return default;
        var spacing = properties.Element(W + "spacing");
        var indent = properties.Element(W + "ind");
        return new PPr(
            NormalizeAlignment(properties.Element(W + "jc")?.Attribute(W + "val")?.Value),
            Points(spacing?.Attribute(W + "before")?.Value),
            Points(spacing?.Attribute(W + "after")?.Value),
            ReadLineSpacing(spacing),
            Chars(indent?.Attribute(W + "firstLineChars")?.Value),
            Points(indent?.Attribute(W + "firstLine")?.Value),
            Chars(indent?.Attribute(W + "hangingChars")?.Value),
            Points(indent?.Attribute(W + "hanging")?.Value));
    }

    private static LineSpacingInfo? ReadLineSpacing(XElement? spacing)
    {
        var lineValue = spacing?.Attribute(W + "line")?.Value;
        if (!double.TryParse(lineValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var line) || line <= 0)
            return null;
        var rule = spacing?.Attribute(W + "lineRule")?.Value ?? "auto";
        return rule switch
        {
            "auto" when Math.Abs(line - 240) < 0.01 => new LineSpacingInfo("single", null, null),
            "auto" => new LineSpacingInfo("multiple", null, line / 240.0),
            "exact" => new LineSpacingInfo("exact", line / 20.0, null),
            "atLeast" => new LineSpacingInfo("atLeast", line / 20.0, null),
            _ => null
        };
    }

    private static bool? ReadBold(XElement? properties)
    {
        var bold = properties?.Element(W + "b");
        if (bold is null) return null;
        var value = bold.Attribute(W + "val")?.Value;
        if (string.IsNullOrWhiteSpace(value)) return true;
        return value is not ("0" or "false" or "off" or "none");
    }

    private static string? NormalizeAlignment(string? value) => value?.Trim().ToLowerInvariant() switch
    {
        "both" or "distribute" or "justify" or "justified" => "justify",
        "start" or "left" => "left",
        "end" or "right" => "right",
        "center" => "center",
        _ => null
    };

    private static double? Chars(string? value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed / 100.0 : null;

    private static double? Points(string? value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed / 20.0 : null;

    private static XDocument ReadXml(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name) ?? throw new InvalidDataException($"Missing OOXML part: {name}");
        using var entryStream = entry.Open();
        return XDocument.Load(entryStream);
    }

    private static XDocument? TryReadXml(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        if (entry is null) return null;
        using var entryStream = entry.Open();
        return XDocument.Load(entryStream);
    }

    private readonly record struct PPr(
        string? Alignment,
        double? BeforePt,
        double? AfterPt,
        LineSpacingInfo? Line,
        double? FirstLineChars,
        double? FirstLinePt,
        double? HangingChars,
        double? HangingPt);

    private sealed record LineSpacingInfo(string Kind, double? Pt, double? Multiple);
    private sealed record StyleDefinition(string? Type, string? BasedOn, PPr Paragraph, bool? Bold);
    private readonly record struct ParagraphResolution(PPr Values, bool IsResolved);
    private readonly record struct BoldResolution(bool? Values, bool IsResolved);
}

