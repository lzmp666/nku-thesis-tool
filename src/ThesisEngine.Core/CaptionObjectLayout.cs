using System.Globalization;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// Conservative layout repair, independent of academic compliance rules. A caption
/// is the visible anchor; only its uniquely adjacent object can be resized. Unknown
/// sections, columns, wrapping, merged cells and inherited table geometry are excluded.
/// </summary>
internal static class CaptionObjectLayout
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    internal static XElement? AdjacentObject(XElement caption, CaptionKind kind)
    {
        var body = caption.Parent;
        if (body?.Name != W + "body") return null;
        var candidates = kind == CaptionKind.Figure ? caption.ElementsBeforeSelf().Reverse() : caption.ElementsAfterSelf();
        foreach (var candidate in candidates)
        {
            if (candidate.Name == W + "p" && !candidate.Descendants(W + "drawing").Any() &&
                !candidate.Descendants(W + "pict").Any() && !candidate.Descendants(W + "sectPr").Any() &&
                !candidate.Descendants(W + "br").Any() && !candidate.Descendants(W + "fldChar").Any() &&
                string.IsNullOrWhiteSpace(string.Concat(candidate.Descendants(W + "t").Select(t => t.Value)))) continue;
            return kind == CaptionKind.Table && candidate.Name == W + "tbl" ||
                kind == CaptionKind.Figure && candidate.Name == W + "p" ? candidate : null;
        }
        return null;
    }

    internal static ObjectWidthSnapshot? Read(XElement caption, CaptionKind kind, XDocument? styles = null)
    {
        var target = AdjacentObject(caption, kind);
        if (target is null) return null;
        var limit = AvailableWidth(target, styles);
        if (limit is null) return null;
        if (kind == CaptionKind.Figure)
        {
            var geometry = PictureGeometry(target);
            if (geometry is null) return null;
            return new ObjectWidthSnapshot(geometry.Value.Width / 12700.0, limit.Value, "picture");
        }
        var columns = TableColumns(target);
        return columns is null ? null : new ObjectWidthSnapshot(columns.Sum() / 20.0, limit.Value, "table");
    }

    private static double? AvailableWidth(XElement target, XDocument? styles)
    {
        // A section's properties occur at its END, including the final body sectPr.
        var section = target.ElementsAfterSelf().Prepend(target).Select(e => e.Name == W + "sectPr" ? e : e.Element(W + "pPr")?.Element(W + "sectPr"))
            .FirstOrDefault(e => e is not null);
        if (section is null) return null;
        var page = section.Element(W + "pgSz");
        var margins = section.Element(W + "pgMar");
        var width = Number(page, W + "w");
        var left = Number(margins, W + "left");
        var right = Number(margins, W + "right");
        var columns = section.Element(W + "cols");
        if (width is null || width <= 0 || width > 31680 || left is null || right is null || left < 0 || right < 0 ||
            (Number(margins, W + "gutter") ?? 0) != 0 ||
            (Number(columns, W + "num") ?? 1) != 1 || columns?.Elements(W + "col").Any() == true) return null;
        var indent = EffectiveIndent(target, styles);
        // Character-based/negative indentation requires Word's font layout.
        if (indent is null || indent.Attributes().Any(a => a.Name.LocalName.EndsWith("Chars", StringComparison.Ordinal))) return null;
        var start = Number(indent, W + "start") ?? Number(indent, W + "left") ?? 0;
        var end = Number(indent, W + "end") ?? Number(indent, W + "right") ?? 0;
        var firstLine = Number(indent, W + "firstLine") ?? 0;
        if (start < 0 || end < 0 || firstLine < 0 || (Number(indent, W + "hanging") ?? 0) != 0) return null;
        var usable = (width.Value - left.Value - right.Value - start - end - firstLine) / 20.0;
        return usable > 0 ? usable : null;
    }

    private static XElement? EffectiveIndent(XElement target, XDocument? styles)
    {
        var indent = new XElement(W + "ind");
        if (target.Name == W + "tbl")
        {
            // No inherited table geometry is silently treated as zero.
            var defaultTableStyle = styles?.Root?.Elements(W + "style").FirstOrDefault(e =>
                (string?)e.Attribute(W + "type") == "table" && (string?)e.Attribute(W + "default") is "1" or "true" or "on");
            if (defaultTableStyle?.Element(W + "basedOn") is not null ||
                defaultTableStyle?.Descendants(W + "tblInd").Any(e => Number(e, W + "w") != 0) == true ||
                defaultTableStyle?.Descendants(W + "tblCellSpacing").Any(e => Number(e, W + "w") != 0) == true) return null;
            return indent;
        }
        void Merge(XElement? properties)
        {
            foreach (var attribute in properties?.Element(W + "ind")?.Attributes() ?? Enumerable.Empty<XAttribute>())
                indent.SetAttributeValue(attribute.Name, attribute.Value);
        }
        Merge(styles?.Root?.Element(W + "docDefaults")?.Element(W + "pPrDefault")?.Element(W + "pPr"));
        var map = styles?.Root?.Elements(W + "style").Where(e => (string?)e.Attribute(W + "type") == "paragraph")
            .ToDictionary(e => (string)e.Attribute(W + "styleId")!, StringComparer.Ordinal) ?? new Dictionary<string, XElement>();
        var id = (string?)target.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") ??
            (string?)map.Values.FirstOrDefault(e => (string?)e.Attribute(W + "default") is "1" or "true" or "on")?.Attribute(W + "styleId");
        var chain = new Stack<XElement>(); var seen = new HashSet<string>(StringComparer.Ordinal);
        while (id is not null)
        {
            if (!seen.Add(id) || !map.TryGetValue(id, out var style)) return null;
            chain.Push(style); id = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
        }
        foreach (var style in chain) Merge(style.Element(W + "pPr"));
        Merge(target.Element(W + "pPr"));
        return indent;
    }

    private static (long Width, long Height)? PictureGeometry(XElement paragraph)
    {
        // No text, equations, multiple objects, floating anchors, charts or SmartArt.
        if (paragraph.Descendants(W + "t").Any(t => !string.IsNullOrWhiteSpace(t.Value)) ||
            paragraph.Descendants(W + "drawing").Count() != 1 || paragraph.Descendants(WP + "anchor").Any() ||
            paragraph.Descendants(W + "pict").Any()) return null;
        var inline = paragraph.Descendants(WP + "inline").SingleOrDefault();
        var picture = inline?.Descendants(PIC + "pic").SingleOrDefault();
        var blip = picture?.Element(PIC + "blipFill")?.Element(A + "blip");
        if (blip?.Attribute(R + "embed") is null || blip.Attribute(R + "link") is not null) return null;
        var effects = inline?.Element(WP + "effectExtent");
        // Word may add vertical glyph/visual padding (e.g. b=6350) after
        // changing a picture size. Only horizontal effects consume line width.
        if ((Number(inline, "distL") ?? 0) != 0 || (Number(inline, "distR") ?? 0) != 0 ||
            effects?.Attributes().Where(a => !a.IsNamespaceDeclaration).Any(a => !long.TryParse(a.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) || value < 0) == true ||
            (Number(effects, "l") ?? 0) != 0 || (Number(effects, "r") ?? 0) != 0) return null;
        var transform = picture?.Element(PIC + "spPr")?.Element(A + "xfrm");
        var extent = inline?.Element(WP + "extent");
        var inner = transform?.Element(A + "ext");
        var width = Number(extent, "cx"); var height = Number(extent, "cy");
        if (width is null || height is null || width <= 0 || height <= 0 || width > 20116800 || height > 20116800 ||
            width != Number(inner, "cx") || height != Number(inner, "cy") ||
            (Number(transform, "rot") ?? 0) != 0) return null;
        return (width.Value, height.Value);
    }

    private static long[]? TableColumns(XElement table)
    {
        var properties = table.Element(W + "tblPr");
        if ((string?)properties?.Element(W + "tblLayout")?.Attribute(W + "type") != "fixed" ||
            properties.Element(W + "tblStyle") is not null || properties.Element(W + "tblpPr") is not null ||
            properties.Element(W + "bidiVisual") is not null ||
            (Number(properties.Element(W + "tblInd"), W + "w") ?? 0) != 0 ||
            (Number(properties.Element(W + "tblCellSpacing"), W + "w") ?? 0) != 0 ||
            table.Descendants(W + "tbl").Any() || table.Descendants(W + "gridSpan").Any() ||
            table.Descendants(W + "vMerge").Any() || table.Descendants(W + "hMerge").Any() ||
            table.Descendants(W + "gridBefore").Any() || table.Descendants(W + "gridAfter").Any() || table.Descendants(W + "tblPrEx").Any() ||
            table.Descendants().Any(e => e.Name.LocalName is "ins" or "del" or "tblPrChange" or "tcPrChange")) return null;
        var columns = table.Element(W + "tblGrid")?.Elements(W + "gridCol").Select(e => Number(e, W + "w") ?? 0).ToArray();
        if (columns is null || columns.Length == 0 || columns.Any(w => w <= 0 || w > 31680) || columns.Sum() > 31680) return null;
        var preferred = properties.Element(W + "tblW");
        if ((string?)preferred?.Attribute(W + "type") != "dxa" || Number(preferred, W + "w") != columns.Sum()) return null;
        var rows = table.Elements(W + "tr").ToArray();
        if (rows.Length == 0) return null;
        foreach (var row in rows)
        {
            var cells = row.Elements(W + "tc").ToArray();
            if (cells.Length != columns.Length) return null;
            for (var i = 0; i < cells.Length; i++)
            {
                var cellWidth = cells[i].Element(W + "tcPr")?.Element(W + "tcW");
                if ((string?)cellWidth?.Attribute(W + "type") != "dxa" || Number(cellWidth, W + "w") != columns[i]) return null;
            }
        }
        return columns;
    }

    internal static void Apply(XElement caption, CaptionKind kind, double targetPoints, XDocument? styles = null)
    {
        var current = Read(caption, kind, styles) ?? throw new InvalidDataException("对象或分节布局不再满足安全缩放条件。");
        if (Math.Abs(current.MaximumWidthPt - targetPoints) > 0.01 || current.WidthPt <= targetPoints)
            throw new InvalidDataException("对象目标宽度与当前页面可用宽度不一致，已停止修改。");
        var target = AdjacentObject(caption, kind)!;
        if (kind == CaptionKind.Figure)
        {
            var geometry = PictureGeometry(target)!.Value;
            var width = (long)Math.Round(targetPoints * 12700);
            var height = (long)Math.Round(geometry.Height * (double)width / geometry.Width);
            var inline = target.Descendants(WP + "inline").Single();
            foreach (var extent in new[] { inline.Element(WP + "extent")!, inline.Descendants(PIC + "pic").Single().Element(PIC + "spPr")!.Element(A + "xfrm")!.Element(A + "ext")! })
            {
                extent.SetAttributeValue("cx", width); extent.SetAttributeValue("cy", height);
            }
        }
        else
        {
            var before = TableColumns(target)!;
            var total = (long)Math.Round(targetPoints * 20);
            var after = before.Select(w => (long)Math.Floor(w * (double)total / before.Sum())).ToArray();
            after[after.Length - 1] += total - after.Sum();
            if (after.Any(w => w <= 0)) throw new InvalidDataException("表格列宽过小，不能安全缩放。");
            var grid = target.Element(W + "tblGrid")!.Elements(W + "gridCol").ToArray();
            for (var i = 0; i < grid.Length; i++) grid[i].SetAttributeValue(W + "w", after[i]);
            foreach (var row in target.Elements(W + "tr"))
                foreach (var pair in row.Elements(W + "tc").Zip(after))
                    pair.First.Element(W + "tcPr")!.Element(W + "tcW")!.SetAttributeValue(W + "w", pair.Second);
            target.Element(W + "tblPr")!.Element(W + "tblW")!.SetAttributeValue(W + "w", total);
        }
    }

    private static long? Number(XElement? element, XName name) =>
        long.TryParse((string?)element?.Attribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : null;
}
