using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace ThesisEngine.Core.Analysis;

/// <summary>估算每个 OOXML 段落所在的页码（1-based）。</summary>
/// <remarks>
/// 这是静态估算：依据节页面高度、上下边距、段落行距/字号/段距、显式分页符和
/// nextPage 分节，按「每行高度 × 行数」累加。它不调用 Word，不做字体度量渲染，
/// 因此与 Word 实机分页可能有 1 页左右偏差；报告必须标注为「估算页码」，
/// 不得当作 Word 实际页码。无法读取页面尺寸时返回 null（未判定）而非猜测。
/// </remarks>
public sealed record ParagraphPageEstimate(
    int ParagraphIndex,
    int? StartPage,
    int? EndPage,
    bool ForcedBreakBefore,
    string Basis);

public sealed record PageMap(
    int? TotalPages,
    IReadOnlyList<ParagraphPageEstimate> Paragraphs,
    string Note);

public sealed class PageNumberEstimator
{
    private const double TwipsPerPoint = 20.0;
    private const double CentimetresPerPoint = 2.54 / 72.0;
    private const double DefaultLineHeightFactor = 1.32; // 单倍行距下每行高度 ≈ 字号 × 1.32
    private const double DefaultFontSizePt = 12.0;
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public PageMap Estimate(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = new FileStream(Path.GetFullPath(path), FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var document = ReadXml(archive, "word/document.xml");
        var probe = new DocxLayoutProbe().Read(path);
        var layout = probe.Paragraphs;
        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidDataException("Missing Word document body.");

        // 显式分页符（段落内 w:br w:type=page）与 pageBreakBefore。
        var forcedBreak = new bool[layout.Count];
        for (var i = 0; i < layout.Count; i++)
        {
            var p = Paragraphs(body).ElementAtOrDefault(i);
            if (p is null) continue;
            forcedBreak[i] = p.Element(W + "pPr")?.Element(W + "pageBreakBefore") is not null
                || p.Descendants(W + "br").Any(b => string.Equals((string?)b.Attribute(W + "type"), "page", StringComparison.Ordinal));
        }

        var xmlParagraphs = Paragraphs(body).ToArray();
        var fontSizes = xmlParagraphs.Select(MaxFontSizePt).ToArray();

        var estimates = new List<ParagraphPageEstimate>(layout.Count);
        var sectionIndex = -1;
        SectionLayout? section = null;
        double pageHeightPt = 0, topPt = 0, bottomPt = 0, usablePt = 0;
        var page = 0;
        var usedPt = 0.0;
        var firstParagraphOnPage = true;

        foreach (var paragraph in layout)
        {
            if (paragraph.SectionIndex != sectionIndex)
            {
                sectionIndex = paragraph.SectionIndex;
                section = probe.Sections.FirstOrDefault(s => s.Index == sectionIndex);
                if (section is not null && section.HeightCm is double h && section.TopCm is double t && section.BottomCm is double b)
                {
                    pageHeightPt = h / CentimetresPerPoint;
                    topPt = t / CentimetresPerPoint;
                    bottomPt = b / CentimetresPerPoint;
                    usablePt = Math.Max(72, pageHeightPt - topPt - bottomPt);
                }
                else
                {
                    usablePt = 0; // 页面尺寸不可读：该节页码标未判定。
                }
                // 分节总是从新页开始（nextPage 是默认；continuous 无法从 OOXML 静态可靠判定，保守换页）。
                page++;
                usedPt = 0;
                firstParagraphOnPage = true;
            }

            if (forcedBreak[paragraph.ParagraphIndex] && !firstParagraphOnPage)
            {
                page++;
                usedPt = 0;
            }

            var canEstimate = usablePt > 0 && paragraph.ValuesKnown;
            if (!canEstimate)
            {
                estimates.Add(new ParagraphPageEstimate(paragraph.ParagraphIndex, null, null,
                    forcedBreak[paragraph.ParagraphIndex], "页面尺寸或段落格式不可读，未判定"));
                firstParagraphOnPage = false;
                continue;
            }

            var fontSize = paragraph.ParagraphIndex >= 0 && paragraph.ParagraphIndex < fontSizes.Length
                ? fontSizes[paragraph.ParagraphIndex]
                : DefaultFontSizePt;
            var lineHeight = LineHeightPt(paragraph, fontSize);
            var text = TextOf(paragraph);
            var charsPerLine = CharsPerLine(section, fontSize);
            var lines = Math.Max(1, (int)Math.Ceiling(text.Length / (double)Math.Max(1, charsPerLine)));
            var blockHeight = paragraph.SpaceBeforePt + paragraph.SpaceAfterPt + lines * lineHeight;

            if (!firstParagraphOnPage && usedPt + blockHeight > usablePt)
            {
                page++;
                usedPt = 0;
            }
            usedPt += blockHeight;
            var startPage = Math.Max(1, page);
            var endPage = startPage;
            if (blockHeight > usablePt)
                endPage = startPage + (int)Math.Floor((blockHeight - usablePt) / usablePt);

            estimates.Add(new ParagraphPageEstimate(paragraph.ParagraphIndex, startPage, endPage,
                forcedBreak[paragraph.ParagraphIndex], "按节页面高度/行距/段距/字号估算"));
            firstParagraphOnPage = false;
        }

        var total = estimates.Count > 0 && estimates.All(e => e.StartPage.HasValue)
            ? estimates.Max(e => e.EndPage)
            : null;
        return new PageMap(total, estimates,
            "估算页码：由 OOXML 页面尺寸、行距、字号、段距与显式分页符推算，未经 Word 实机渲染；与 Word 实际分页可能有偏差，不作为页码合规判据。");
    }

    private static IEnumerable<XElement> Paragraphs(XElement body) =>
        body.Elements().Where(e => e.Name.LocalName is "p" or "tbl")
            .SelectMany(e => e.Name.LocalName == "p" ? new[] { e } : e.Descendants(W + "p"));

    private static string TextOf(ParagraphLayout p)
    {
        var sb = new StringBuilder();
        foreach (var run in p.Runs) sb.Append(run.Text);
        return sb.ToString();
    }

    /// <summary>取段落内显式 w:sz 的最大值（半点 → 磅）；没有则回退 12pt。</summary>
    private static double MaxFontSizePt(XElement paragraph)
    {
        var halves = paragraph.Descendants(W + "sz")
            .Select(s => (string?)s.Attribute(W + "val"))
            .Select(v => int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0)
            .Where(n => n > 0)
            .ToArray();
        return halves.Length == 0 ? DefaultFontSizePt : halves.Max() / 2.0;
    }

    private static double LineHeightPt(ParagraphLayout p, double fontSize)
    {
        if (p.LineSpacingKind == "exact" && p.LineSpacingPt is double exact) return Math.Max(6, exact);
        if (p.LineSpacingKind == "atLeast" && p.LineSpacingPt is double least) return Math.Max(6, least);
        if (p.LineSpacingKind == "multiple" && p.LineSpacingMultiple is double m) return fontSize * DefaultLineHeightFactor * m;
        return fontSize * DefaultLineHeightFactor; // single
    }

    private static int CharsPerLine(SectionLayout? section, double fontSizePt)
    {
        var widthCm = section?.WidthCm ?? 21.0;
        var left = section?.LeftCm ?? 3.2;
        var right = section?.RightCm ?? 3.2;
        var gutter = section?.GutterCm ?? 0;
        var usableCm = Math.Max(4, widthCm - left - right - gutter);
        var charWidthCm = fontSizePt * CentimetresPerPoint; // 汉字近似正方
        return Math.Max(8, (int)Math.Floor(usableCm / Math.Max(0.1, charWidthCm)));
    }

    private static XDocument ReadXml(ZipArchive archive, string entry)
    {
        var item = archive.GetEntry(entry) ?? throw new InvalidDataException($"Missing {entry}.");
        using var stream = item.Open();
        return XDocument.Load(stream);
    }
}
