using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ThesisEngine.Core.Document;

/// <summary>
/// 文档级版式分析器（#112）。
///
/// 现有检查都停在段落级，页面设置一概没读。纸张、页边距、页眉页脚、页码、目录这几项
/// 恰恰是答辩前最容易被退回来的，而且都是 Word 里一次设置全文生效的东西。
///
/// 两条口径贯穿全文：
/// 一、有条款依据才报。每条都带 R1 / R2 / R3 的出处，条款查不到的项列进「未检查」，
///    不拿猜的数字去凑问题数。
/// 二、不假定每一节都该有页眉。规范说「篇眉从中文摘要开始」，封面本来就没有页眉，
///    逐节要求页眉会把每篇合格论文都报错。所以「有没有页眉」按全文判断。
/// </summary>
public sealed class DocumentAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>1 cm 对应的 twips。</summary>
    internal const double TwipsPerCm = 566.9291338582677;

    /// <summary>页面尺寸容差，0.1cm。Word 按整 twips 存储，取整误差远小于这个数。</summary>
    internal const double ToleranceCm = 0.1;

    // R1（4.1 表4.1）规定值
    internal const double MarginTopBottomCm = 3.8;
    internal const double MarginLeftRightCm = 3.2;
    internal const double HeaderFooterCm = 3.0;
    internal const double GutterCm = 0.0;

    // A4 210×297mm
    internal const double A4WidthCm = 21.0;
    internal const double A4HeightCm = 29.7;

    /// <summary>页码域指令。要排除 NUMPAGES——那是总页数，不是页码。</summary>
    private static readonly Regex PageFieldPattern =
        new(@"(?<!\w)PAGE(?!\w)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex TocFieldPattern =
        new(@"(?<!\w)TOC(?!\w)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly string[] TocHeadings = { "目录", "目 录", "Contents", "TABLE OF CONTENTS" };

    public DocumentAnalysisResult Analyze(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var document = LoadXml(archive, "word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");
        var settings = LoadXml(archive, "word/settings.xml");
        var styles = LoadXml(archive, "word/styles.xml");
        var relationships = ReadRelationships(archive);

        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidDataException("文档中没有 body");

        var sectionElements = CollectSectionProperties(body);
        var evenAndOdd = IsSettingOn(settings, "evenAndOddHeaders");

        var sections = new List<SectionInfo>();
        var issues = new List<DocumentIssue>();

        for (var i = 0; i < sectionElements.Count; i++)
        {
            var sectPr = sectionElements[i];
            var info = ReadSection(i + 1, sectPr, archive, relationships);
            sections.Add(info);

            CheckPageSize(info, issues);
            CheckMargins(info, issues);
            CheckHeaderFooterDistance(info, issues);
            CheckGutter(info, issues);
            CheckPageNumberFormatValue(info, issues);
        }

        CheckWholeDocument(sections, sectionElements, archive, relationships, evenAndOdd, body, issues);

        // R2 / R3 的字体字号与居中。样式链展开不了的部分不报问题，单列「未判定」。
        var formatResult = CheckHeaderFooterFormats(sectionElements, archive, relationships, styles, issues);

        var uncheckedItems = UncheckedItems
            .Concat(formatResult.Undetermined)
            .ToList();

        return new DocumentAnalysisResult(
            sections,
            issues,
            HasTocField: HasField(body, TocFieldPattern),
            EvenAndOddHeaders: evenAndOdd,
            UpdateFieldsOnOpen: IsSettingOn(settings, "updateFields"),
            PartReadings: formatResult.Readings,
            UncheckedItems: uncheckedItems);
    }

    /// <summary>
    /// 按文档顺序取所有 sectPr。分节符写在段落属性里，最后一节写在 body 末尾。
    /// 顺序错了后面每一项都会对应到错误的节上，所以单独一步。
    /// </summary>
    internal static List<XElement> CollectSectionProperties(XElement body)
    {
        var result = new List<XElement>();
        foreach (var paragraph in body.Descendants(W + "p"))
        {
            var sectPr = paragraph.Element(W + "pPr")?.Element(W + "sectPr");
            if (sectPr != null) result.Add(sectPr);
        }

        var bodySectPr = body.Element(W + "sectPr");
        if (bodySectPr != null)
        {
            result.Add(bodySectPr);
        }
        else if (result.Count > 0)
        {
            // 没有 body 级 sectPr 时，最后一节的属性就是文档最终的页面设置。
            // 不补这一条会少一节，页面检查就漏了。
            result.Add(result[^1]);
        }

        return result;
    }

    private static SectionInfo ReadSection(
        int index,
        XElement sectPr,
        ZipArchive archive,
        IReadOnlyDictionary<string, string> relationships)
    {
        var pageSize = sectPr.Element(W + "pgSz");
        var margin = sectPr.Element(W + "pgMar");
        var numberType = sectPr.Element(W + "pgNumType");

        return new SectionInfo(
            Index: index,
            PageWidthCm: ToCm(pageSize?.Attribute(W + "w")?.Value),
            PageHeightCm: ToCm(pageSize?.Attribute(W + "h")?.Value),
            TopCm: ToCm(margin?.Attribute(W + "top")?.Value),
            BottomCm: ToCm(margin?.Attribute(W + "bottom")?.Value),
            LeftCm: ToCm(margin?.Attribute(W + "left")?.Value),
            RightCm: ToCm(margin?.Attribute(W + "right")?.Value),
            HeaderCm: ToCm(margin?.Attribute(W + "header")?.Value),
            FooterCm: ToCm(margin?.Attribute(W + "footer")?.Value),
            GutterCm: ToCm(margin?.Attribute(W + "gutter")?.Value),
            PageNumberFormat: numberType?.Attribute(W + "fmt")?.Value,
            PageNumberStart: ParseInt(numberType?.Attribute(W + "start")?.Value),
            TitlePage: sectPr.Element(W + "titlePg") != null,
            HasHeaderContent: HasPartContent(sectPr, archive, relationships, "headerReference"),
            HasFooterContent: HasPartContent(sectPr, archive, relationships, "footerReference"),
            HeaderCount: sectPr.Elements(W + "headerReference").Count(),
            FooterCount: sectPr.Elements(W + "footerReference").Count());
    }

    private static void CheckPageSize(SectionInfo section, List<DocumentIssue> issues)
    {
        if (section.PageWidthCm <= 0 || section.PageHeightCm <= 0)
        {
            issues.Add(new DocumentIssue(
                Code: "D1",
                Title: "未读到纸张尺寸",
                Basis: "R1（4.1 表4.1）：纸张 A4",
                SectionIndex: section.Index,
                Current: "sectPr 里没有 pgSz",
                Expected: "A4（21.0 × 29.7 cm）"));
            return;
        }

        var isA4 = Math.Abs(section.PageWidthCm - A4WidthCm) <= ToleranceCm &&
                   Math.Abs(section.PageHeightCm - A4HeightCm) <= ToleranceCm;
        if (isA4) return;

        issues.Add(new DocumentIssue(
            Code: "D1",
            Title: "纸张尺寸不是 A4",
            Basis: "R1（4.1 表4.1）：纸张 A4（210×297）",
            SectionIndex: section.Index,
            Current: $"{section.PageWidthCm:0.##} × {section.PageHeightCm:0.##} cm",
            Expected: $"{A4WidthCm:0.#} × {A4HeightCm:0.#} cm"));
    }

    private static void CheckMargins(SectionInfo section, List<DocumentIssue> issues)
    {
        CheckLength(issues, "D2", "上页边距", section.Index, section.TopCm, MarginTopBottomCm, "R1（4.1 表4.1）：上 3.8cm");
        CheckLength(issues, "D2", "下页边距", section.Index, section.BottomCm, MarginTopBottomCm, "R1（4.1 表4.1）：下 3.8cm");
        CheckLength(issues, "D2", "左页边距", section.Index, section.LeftCm, MarginLeftRightCm, "R1（4.1 表4.1）：左 3.2cm");
        CheckLength(issues, "D2", "右页边距", section.Index, section.RightCm, MarginLeftRightCm, "R1（4.1 表4.1）：右 3.2cm");
    }

    private static void CheckHeaderFooterDistance(SectionInfo section, List<DocumentIssue> issues)
    {
        CheckLength(issues, "D3", "页眉距边界", section.Index, section.HeaderCm, HeaderFooterCm, "R1（4.1 表4.1）：页眉 3.0cm");
        CheckLength(issues, "D3", "页脚距边界", section.Index, section.FooterCm, HeaderFooterCm, "R1（4.1 表4.1）：页脚 3.0cm");
    }

    private static void CheckGutter(SectionInfo section, List<DocumentIssue> issues)
    {
        CheckLength(issues, "D4", "装订线", section.Index, section.GutterCm, GutterCm, "R1（4.1 表4.1）：装订线 0cm");
    }

    private static void CheckLength(
        List<DocumentIssue> issues,
        string code,
        string name,
        int sectionIndex,
        double actualCm,
        double expectedCm,
        string basis)
    {
        if (Math.Abs(actualCm - expectedCm) <= ToleranceCm) return;

        issues.Add(new DocumentIssue(
            Code: code,
            Title: $"{name}不符合规范",
            Basis: basis,
            SectionIndex: sectionIndex,
            Current: $"{actualCm:0.##} cm",
            Expected: $"{expectedCm:0.#} cm"));
    }

    /// <summary>
    /// 页眉：只有「所有节都没有页眉内容」才算问题。
    ///
    /// 这条必须按全文判断。实测 Word 保存文档时会为每一节都建好 default / even / first
    /// 三组页眉页脚引用，即使内容是空的（见 tmp 里 Word 16 生成的样本）。而规范说的是
    /// 「篇眉从中文摘要开始」，封面本来就没有页眉。逐节要求页眉，等于把每篇合格论文都报一遍。
    /// 「哪一节是封面」在 XML 层面判不出来，所以只报全文缺失，局部为空放在读数里让用户自己看。
    /// </summary>
    private static void CheckHeaders(
        IReadOnlyList<SectionInfo> sections,
        ZipArchive archive,
        IReadOnlyDictionary<string, string> relationships,
        List<DocumentIssue> issues)
    {
        if (sections.Any(section => section.HasHeaderContent)) return;

        var withReference = sections.Where(section => section.HeaderCount > 0).Select(section => section.Index).ToList();
        var detail = withReference.Count == 0
            ? "没有任何节设置页眉"
            : $"第 {string.Join("、", withReference)} 节引用了页眉部件，但里面没有文字";

        issues.Add(new DocumentIssue(
            Code: "D6",
            Title: "全文没有页眉内容",
            Basis: "R2（3.4）：篇眉从中文摘要开始，内容与该部分的标题相同",
            SectionIndex: 0,
            Current: detail,
            Expected: "从中文摘要所在部分开始设置页眉，内容用该部分的标题"));
    }

    /// <summary>
    /// 页码：全文所有页脚里都没有 PAGE 域才报。手打的数字不算——
    /// 它不会随页面变化，正文一改页码就全错。
    /// </summary>
    private static void CheckPageNumberField(
        IReadOnlyList<SectionInfo> sections,
        IReadOnlyList<XElement> sectionElements,
        ZipArchive archive,
        IReadOnlyDictionary<string, string> relationships,
        List<DocumentIssue> issues)
    {
        var footerParts = sectionElements
            .SelectMany(sectPr => sectPr.Elements(W + "footerReference"))
            .Select(reference => ResolvePartName(reference, relationships))
            .Where(name => name != null)
            .Select(name => LoadXml(archive, name!))
            .Where(part => part != null)
            .ToList();

        if (footerParts.Any(part => HasField(part!.Root!, PageFieldPattern))) return;

        var withFooter = sections.Where(section => section.FooterCount > 0).Select(section => section.Index).ToList();
        var detail = withFooter.Count == 0
            ? "所有节都没有页脚"
            : $"第 {string.Join("、", withFooter)} 节的页脚里没有 PAGE 域（手打数字不算）";

        issues.Add(new DocumentIssue(
            Code: "D7",
            Title: "全文没有页码域",
            Basis: "R3（4.1 表4.1）：页码从第一章开始按阿拉伯数字连续编排",
            SectionIndex: 0,
            Current: detail,
            Expected: "在页脚插入 PAGE 域，前置部分用大写罗马数字、正文用阿拉伯数字"));
    }

    private static void CheckPageNumberFormatValue(SectionInfo section, List<DocumentIssue> issues)
    {
        var format = section.PageNumberFormat;
        if (format == null) return;
        if (format is "decimal" or "upperRoman" or "lowerRoman") return;

        issues.Add(new DocumentIssue(
            Code: "D8",
            Title: "页码编号格式不是阿拉伯或罗马数字",
            Basis: "R3（4.1 表4.1）：正文阿拉伯数字、前置部分大写罗马数字",
            SectionIndex: section.Index,
            Current: format,
            Expected: "decimal 或 upperRoman"));
    }

    /// <summary>
    /// 全文级检查。这几项判断的是「整篇有没有这个东西」，不是逐节，
    /// 因为规范对页眉的起点有明确规定（中文摘要开始），逐节要求必然误报。
    /// </summary>
    private static void CheckWholeDocument(
        IReadOnlyList<SectionInfo> sections,
        IReadOnlyList<XElement> sectionElements,
        ZipArchive archive,
        IReadOnlyDictionary<string, string> relationships,
        bool evenAndOdd,
        XElement body,
        List<DocumentIssue> issues)
    {
        CheckHeaders(sections, archive, relationships, issues);
        CheckPageNumberField(sections, sectionElements, archive, relationships, issues);

        // 页码格式是否区分了前置与正文。只在多节且格式完全一致时提示，
        // 单节文档判不出分界，硬判就是瞎猜。
        if (sections.Count >= 2)
        {
            var formats = sections.Select(section => section.PageNumberFormat ?? "未设置").Distinct().ToList();
            if (formats.Count == 1)
            {
                issues.Add(new DocumentIssue(
                    Code: "D8",
                    Title: "各节页码格式相同，未区分前置部分与正文",
                    Basis: "R3（4.1 表4.1）：前置部分用大写罗马数字，正文从第一章开始用阿拉伯数字",
                    SectionIndex: 0,
                    Current: $"{sections.Count} 个节的页码格式都是「{formats[0]}」",
                    Expected: "前置部分 upperRoman，正文 decimal"));
            }
        }

        if (evenAndOdd)
        {
            issues.Add(new DocumentIssue(
                Code: "D5",
                Title: "开启了奇偶页不同",
                Basis: "R2（3.4）：论文页眉奇偶页相同",
                SectionIndex: 0,
                Current: "设置了奇偶页不同（w:evenAndOddHeaders）",
                Expected: "关闭奇偶页不同，奇偶页使用同一页眉"));
        }

        CheckTableOfContents(body, issues);
        CheckBlankPages(body, issues);
    }

    /// <summary>
    /// 目录是不是自动生成的域。手打的目录在 Word 里页码对不上，改一次正文就全废。
    /// </summary>
    private static void CheckTableOfContents(XElement body, List<DocumentIssue> issues)
    {
        if (HasField(body, TocFieldPattern)) return;

        var hasTocHeading = body.Descendants(W + "p")
            .Any(para =>
            {
                var text = string.Concat(para.Descendants(W + "t").Select(t => t.Value)).Trim();
                return TocHeadings.Contains(text, StringComparer.Ordinal);
            });

        if (!hasTocHeading) return;

        issues.Add(new DocumentIssue(
            Code: "D9",
            Title: "目录不是自动生成的域",
            Basis: "R6（4.6 表4.5）：目录由 Word 自动生成并随正文更新",
            SectionIndex: 0,
            Current: "文档里有「目录」标题，但没有 TOC 域（多半是手打的）",
            Expected: "用引用 → 目录插入自动目录，正文改动后右键更新域"));
    }

    /// <summary>
    /// 疑似空白页。只报 XML 层面能确定的情形，段后间距、表格跨页要真实分页才判得准，
    /// 不靠推测凑数。
    ///
    /// Word 16 双页视图实测修正了一处口径：单个空段落带分页符只是「多出一页」，
    /// 第二页仍有正文，并不存在整页空白。真正造成整页空白的是连续两个以上的
    /// 空段落各带分页符，所以把这两种分开报，措辞也不混用。
    /// </summary>
    private static void CheckBlankPages(XElement body, List<DocumentIssue> issues)
    {
        var paragraphs = body.Descendants(W + "p").ToList();

        for (var i = 0; i < paragraphs.Count; i++)
        {
            var paragraph = paragraphs[i];

            if (IsEmptyParagraph(paragraph) && HasPageBreak(paragraph))
            {
                var last = i;
                while (last + 1 < paragraphs.Count &&
                       IsEmptyParagraph(paragraphs[last + 1]) &&
                       HasPageBreak(paragraphs[last + 1]))
                {
                    last++;
                }

                if (last > i)
                {
                    issues.Add(new DocumentIssue(
                        Code: "D10",
                        Title: "连续分页符造成整页空白",
                        Basis: "排版惯例（规范未规定空白页，本条不计入规范符合性）",
                        SectionIndex: 0,
                        Current: $"第 {i + 1}–{last + 1} 段是连续 {last - i + 1} 个空段落，各带一个分页符",
                        Expected: "只留一个分页符，删掉多余的空段落（Word 实测这种写法会多出一整页空白）"));
                    i = last;
                    continue;
                }

                issues.Add(new DocumentIssue(
                    Code: "D10",
                    Title: "疑似空白页：空段落里带分页符",
                    Basis: "排版惯例（规范未规定空白页，本条不计入规范符合性）",
                    SectionIndex: 0,
                    Current: $"第 {i + 1} 段是空段落，内含分页符；Word 实测只会多一页，第二页仍有正文，不等于整页空白",
                    Expected: "确认是否真要另起一页；要的话把分页放到起新页的段落上并删掉空段落"));
            }

            var pageBreakBefore = paragraph.Element(W + "pPr")?.Element(W + "pageBreakBefore") != null;
            if (!pageBreakBefore || i == 0) continue;

            if (!IsEmptyParagraph(paragraphs[i - 1])) continue;

            issues.Add(new DocumentIssue(
                Code: "D10",
                Title: "疑似空白页：空段落后面紧跟段前分页",
                Basis: "排版惯例（规范未规定空白页，本条不计入规范符合性）",
                SectionIndex: 0,
                Current: $"第 {i} 段为空，第 {i + 1} 段设置了段前分页",
                Expected: "删掉空段落，只保留段前分页"));
        }
    }

    private static bool IsEmptyParagraph(XElement paragraph) =>
        string.IsNullOrWhiteSpace(
            string.Concat(paragraph.Descendants(W + "t").Select(node => node.Value)));

    private static bool HasPageBreak(XElement paragraph) =>
        paragraph.Descendants(W + "br")
            .Any(br => br.Attribute(W + "type")?.Value == "page");

    /// <summary>
    /// R2 篇眉、R3 页码的字体字号与居中。
    ///
    /// 两处刻意收窄口径：
    /// 一、页脚只判含 PAGE 域的段落。R3 管的是页码，页脚里放校名、密级之类
    ///     的段落不归它管，硬判会报一堆改不掉的问题。
    /// 二、全西文的篇眉按 R2「Abstract 部分用 Times New Roman」判；含中文的按
    ///     「宋体」判。XML 层面认不出哪一节是 Abstract，只能按文字近似，这条
    ///     在「未检查」里也写明了。
    /// </summary>
    private static PartFormatResult CheckHeaderFooterFormats(
        IReadOnlyList<XElement> sectionElements,
        ZipArchive archive,
        IReadOnlyDictionary<string, string> relationships,
        XDocument? styles,
        List<DocumentIssue> issues)
    {
        var allReadings = new List<PartFormatReading>();
        var undetermined = new List<string>();

        foreach (var (name, part) in ResolveParts(sectionElements, archive, relationships, "headerReference"))
        {
            var result = PartFormatReader.Read(part, name, styles);
            undetermined.AddRange(result.Undetermined);
            allReadings.AddRange(result.Readings);
            DescribeUndetermined(result.Readings, "篇眉", undetermined);

            foreach (var reading in result.Readings)
            {
                CheckPartFormat(reading, "D11", "篇眉", westernFontByClause: true, issues);
            }
        }

        foreach (var (name, part) in ResolveParts(sectionElements, archive, relationships, "footerReference"))
        {
            var paragraphs = part.Root?.Descendants(W + "p").ToList() ?? new List<XElement>();
            var result = PartFormatReader.Read(part, name, styles);
            undetermined.AddRange(result.Undetermined);
            allReadings.AddRange(result.Readings);

            var pageFieldParagraphs = result.Readings
                .Where(reading => reading.ParagraphIndex <= paragraphs.Count &&
                                  HasField(paragraphs[reading.ParagraphIndex - 1], PageFieldPattern))
                .ToList();

            var otherCount = result.Readings.Count - pageFieldParagraphs.Count;
            if (otherCount > 0)
            {
                undetermined.Add(
                    $"{name}：{otherCount} 个文字段落不含 PAGE 域，R3 只规定页码，这些段落的字体字号未判定");
            }

            foreach (var reading in pageFieldParagraphs)
            {
                CheckPartFormat(reading, "D12", "页码", westernFontByClause: false, issues);
            }
            DescribeUndetermined(pageFieldParagraphs, "页码", undetermined);
        }

        return new PartFormatResult(allReadings, undetermined);
    }

    /// <summary>
    /// 判不出来的维度逐条写进「未检查」。报告里 0 条问题必须能看出
    /// 是「查了没问题」还是「这一项压根没查」。
    /// </summary>
    private static void DescribeUndetermined(
        IEnumerable<PartFormatReading> readings,
        string subject,
        List<string> undetermined)
    {
        foreach (var reading in readings)
        {
            var missing = new List<string>();
            if (!reading.FontResolved) missing.Add("字体（rPr 与样式链里都没有 rFonts）");
            if (!reading.SizeResolved) missing.Add("字号（rPr 与样式链里都没有 w:sz）");
            if (!reading.AlignmentResolved) missing.Add("对齐（没写 jc 且样式链不可用）");
            if (missing.Count == 0) continue;

            undetermined.Add(
                $"{reading.PartName} 第 {reading.ParagraphIndex} 段：{subject}的" +
                $"{string.Join("、", missing)}未判定，不计入 0 条问题");
        }
    }

    private static void CheckPartFormat(
        PartFormatReading reading,
        string code,
        string subject,
        bool westernFontByClause,
        List<DocumentIssue> issues)
    {
        var location = $"{reading.PartName} 第 {reading.ParagraphIndex} 段「{TrimForReport(reading.Text)}」";
        var isAbstract = westernFontByClause && reading.WesternOnly;

        var fontBasis = westernFontByClause
            ? "R2（4.1 表4.1）：篇眉宋体 10.5 磅居中；Abstract 部分用 Times New Roman 10.5 磅"
            : "R3（4.1 表4.1）：页码宋体 10.5 磅居中";
        var sizeBasis = westernFontByClause
            ? "R2（4.1 表4.1）：篇眉 10.5 磅（五号）"
            : "R3（4.1 表4.1）：页码 10.5 磅（五号）";
        var alignmentBasis = westernFontByClause
            ? "R2（4.1 表4.1）：篇眉居中"
            : "R3（4.1 表4.1）：页码居中";

        if (reading.FontResolved)
        {
            var expectedFont = isAbstract ? "Times New Roman" : "宋体";
            var actualFont = isAbstract ? reading.WesternFont : reading.EastAsiaFont;
            var matched = isAbstract
                ? PartFormatReader.IsTimesNewRoman(actualFont)
                : PartFormatReader.IsSimSun(actualFont);

            if (!matched)
            {
                issues.Add(new DocumentIssue(
                    Code: code,
                    Title: $"{subject}字体不符合规范",
                    Basis: fontBasis,
                    SectionIndex: 0,
                    Current: $"{location} 当前 {actualFont ?? "未设置"}",
                    Expected: expectedFont));
            }
        }

        if (reading.SizeResolved && reading.SizeHalfPoints != PartFormatReader.RequiredSizeHalfPoints)
        {
            issues.Add(new DocumentIssue(
                Code: code,
                Title: $"{subject}字号不符合规范",
                Basis: sizeBasis,
                SectionIndex: 0,
                Current: $"{location} 当前 {reading.SizeHalfPoints / 2.0:0.#} 磅（sz={reading.SizeHalfPoints}）",
                Expected: "10.5 磅（sz=21）"));
        }

        if (reading.AlignmentResolved &&
            !string.Equals(reading.Alignment, "center", StringComparison.OrdinalIgnoreCase))
        {
            issues.Add(new DocumentIssue(
                Code: code,
                Title: $"{subject}未居中",
                Basis: alignmentBasis,
                SectionIndex: 0,
                Current: $"{location} 当前对齐方式 {reading.Alignment}",
                Expected: "居中（w:jc=center）"));
        }
    }

    private static string TrimForReport(string text) =>
        text.Length <= 20 ? text : text[..20] + "…";

    /// <summary>按部件名去重：多个节引用同一个页眉部件，不该读三遍报三遍。</summary>
    private static List<(string Name, XDocument Part)> ResolveParts(
        IReadOnlyList<XElement> sectionElements,
        ZipArchive archive,
        IReadOnlyDictionary<string, string> relationships,
        string referenceName)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<(string, XDocument)>();

        foreach (var reference in sectionElements.SelectMany(sectPr => sectPr.Elements(W + referenceName)))
        {
            var name = ResolvePartName(reference, relationships);
            if (name is null || !seen.Add(name)) continue;

            var part = LoadXml(archive, name);
            if (part is not null) result.Add((name, part));
        }

        return result;
    }

    private static bool HasPartContent(
        XElement sectPr,
        ZipArchive archive,
        IReadOnlyDictionary<string, string> relationships,
        string referenceName)
    {
        return sectPr.Elements(W + referenceName)
            .Select(reference => ResolvePartName(reference, relationships))
            .Where(name => name != null)
            .Select(name => LoadXml(archive, name!))
            .Where(part => part != null)
            .Any(part => part!.Root!.Descendants(W + "t")
                .Any(t => !string.IsNullOrWhiteSpace(t.Value)));
    }

    /// <summary>域：既认 fldSimple 的 instr 属性，也认 fldChar + instrText 三段式。</summary>
    private static bool HasField(XElement scope, Regex pattern)
    {
        foreach (var simple in scope.Descendants(W + "fldSimple"))
        {
            var instruction = simple.Attribute(W + "instr")?.Value;
            if (instruction != null && pattern.IsMatch(instruction)) return true;
        }

        return scope.Descendants(W + "instrText").Any(text => pattern.IsMatch(text.Value));
    }

    private static string? ResolvePartName(
        XElement reference,
        IReadOnlyDictionary<string, string> relationships)
    {
        var id = reference.Attribute(R + "id")?.Value;
        if (id == null) return null;
        return relationships.TryGetValue(id, out var target) ? NormalizePartName(target) : null;
    }

    private static string NormalizePartName(string target)
    {
        var cleaned = target.Replace('\\', '/');
        return cleaned.StartsWith('/') ? cleaned.TrimStart('/') : "word/" + cleaned;
    }

    private static Dictionary<string, string> ReadRelationships(ZipArchive archive)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var entry = archive.GetEntry("word/_rels/document.xml.rels");
        if (entry == null) return result;

        using var reader = new StreamReader(entry.Open());
        var doc = XDocument.Load(reader);
        foreach (var relationship in doc.Root?.Elements(Rel + "Relationship") ?? Enumerable.Empty<XElement>())
        {
            var id = relationship.Attribute("Id")?.Value;
            var target = relationship.Attribute("Target")?.Value;
            if (id != null && target != null) result[id] = target;
        }

        return result;
    }

    private static XDocument? LoadXml(ZipArchive archive, string partName)
    {
        var entry = archive.GetEntry(partName);
        if (entry == null) return null;
        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }

    private static bool IsSettingOn(XDocument? settings, string name)
    {
        var element = settings?.Root?.Element(W + name);
        if (element == null) return false;
        var value = element.Attribute(W + "val")?.Value;
        return value is null or "1" or "true" or "on";
    }

    private static double ToCm(string? twips)
    {
        if (!double.TryParse(twips, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            return 0;
        return Math.Round(value / TwipsPerCm, 3);
    }

    private static int? ParseInt(string? value) =>
        int.TryParse(value, out var parsed) ? parsed : null;

    private static readonly IReadOnlyList<string> UncheckedItems = new[]
    {
        "某一节页眉为空：多节文档的封面节本来就没有页眉，规范只说「从中文摘要开始」，判不出哪节是封面，只报全文缺失",
        "页眉横线：R2/R3 只规定了字体字号与居中，没有规定有无横线，没有条款依据，不查",
        "页眉是否为 Abstract 部分：XML 层面认不出章节归属，只能按文字是否全为西文近似，判错的可能要人工确认",
        "同一页眉/页脚段落里多个 run 格式不一致：取哪个值都不代表这一段，标未判定，需人工看",
        "正文从第一章起编阿拉伯数字：要判断哪一节是「第一章」，靠 XML 判不准，需人工确认",
        "奇偶页是否真的「内容相同」：只查了是否开启奇偶页不同，未比对奇偶页眉文字",
        "空白页的完整判定：段后间距、表格跨页导致的空白要真实分页才算得准，只报两种确定性情形",
        "页面方向与装订线位置：规范未作规定，不查",
    };
}

public sealed record DocumentAnalysisResult(
    IReadOnlyList<SectionInfo> Sections,
    IReadOnlyList<DocumentIssue> Issues,
    bool HasTocField,
    bool EvenAndOddHeaders,
    bool UpdateFieldsOnOpen,
    IReadOnlyList<PartFormatReading> PartReadings,
    IReadOnlyList<string> UncheckedItems)
{
    public IEnumerable<DocumentIssue> ClauseIssues => Issues.Where(issue => issue.HasClause);

    public IEnumerable<DocumentIssue> ConventionIssues => Issues.Where(issue => !issue.HasClause);
}

public sealed record SectionInfo(
    int Index,
    double PageWidthCm,
    double PageHeightCm,
    double TopCm,
    double BottomCm,
    double LeftCm,
    double RightCm,
    double HeaderCm,
    double FooterCm,
    double GutterCm,
    string? PageNumberFormat,
    int? PageNumberStart,
    bool TitlePage,
    bool HasHeaderContent,
    bool HasFooterContent,
    int HeaderCount,
    int FooterCount);

public sealed record DocumentIssue(
    string Code,
    string Title,
    string Basis,
    int SectionIndex,
    string Current,
    string Expected)
{
    /// <summary>惯例项不算进规范符合性，报告里分开列。</summary>
    public bool HasClause => !Basis.StartsWith("排版惯例", StringComparison.Ordinal);
}
