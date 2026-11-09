using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core.Formula;

/// <summary>
/// 公式格式写入（#114 第二档：只改格式，不碰内容）。
///
/// 硬边界，改代码前先读这三条：
/// 1. <c>m:t</c> 里的公式文本一个字符都不许改——这里的所有写法都是「加属性 / 加元素」，
///    没有一处去读改 <c>m:t</c> 的值。改完由调用方逐字比对，不靠眼看。
/// 2. 行内公式整段不碰：它只是句子的一部分，居中、右对齐对它没有意义。
/// 3. 式号缺圆括号不修。那是改文字，不在「四类」里；本类只报出来让人工补。
///
/// 识别逻辑不重新写一遍，一律走 <see cref="FormulaAnalyzer"/> 的
/// <c>DetermineLayout</c> / <c>ReadTextWidthTwips</c>：检查和修复必须认同一批公式，
/// 否则会出现「查出 5 条、只修了 3 条、剩下 2 条再也找不到」的情况（#118 踩过）。
/// </summary>
public sealed class FormulaFormatter
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>把公式排到版心右缘；容差沿用检查器，免得修完仍被判定为「不够右」。</summary>
    private const int RightTabToleranceTwips = FormulaAnalyzer.TabStopToleranceTwips;

    public FormulaFormatResult Apply(string inputPath, string outputPath)
    {
        var doc = ReadDocumentXml(inputPath);
        var body = doc.Root?.Element(W + "body")
            ?? throw new InvalidDataException("文档中没有 body");

        var textWidth = FormulaAnalyzer.ReadTextWidthTwips(body);
        var paragraphs = body.Descendants(W + "p").ToList();

        var fixes = new List<FormulaFix>();
        var skips = new List<string>();
        var standalone = 0;
        var inline = 0;
        var tableRow = 0;

        for (var i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var paragraphIndex = i + 1;

            var hasMath = para.Descendants(M + "oMath")
                .Any(math => !math.Ancestors(M + "oMath").Any());
            if (!hasMath) continue;

            var layout = FormulaAnalyzer.DetermineLayout(
                para, out var numberText, out var numberRun, out var numberInMath);

            if (layout == FormulaLayout.Inline)
            {
                inline++;
                continue;
            }

            if (layout == FormulaLayout.TableRow)
            {
                tableRow++;
                skips.Add($"段落 {paragraphIndex}：表格排版的公式不自动修（位置由列宽决定，改 jc 看不出效果），需人工确认");
                continue;
            }

            standalone++;
            var formulaIndex = standalone;

            FixCentering(para, paragraphIndex, formulaIndex, fixes);
            FixSpacing(para, paragraphIndex, formulaIndex, fixes);
            FixMathFont(para, paragraphIndex, formulaIndex, fixes);
            FixNumberAlignment(para, numberText, numberRun, numberInMath, textWidth, paragraphIndex, formulaIndex, fixes, skips);
            FixNumberFormat(numberText, numberRun, numberInMath, paragraphIndex, formulaIndex, fixes, skips);
        }

        WriteDocumentXml(inputPath, outputPath, doc);

        return new FormulaFormatResult(standalone, inline, tableRow, fixes, skips);
    }

    /// <summary>F1：独立公式居中。段落 jc 和 oMathParaPr/jc 都设，两处说法不一致时 Word 会各按各的显示。</summary>
    private static void FixCentering(XElement para, int paragraphIndex, int formulaIndex, List<FormulaFix> fixes)
    {
        var pPr = GetOrAddPPr(para);
        var paragraphAlign = pPr.Element(W + "jc")?.Attribute(W + "val")?.Value;

        var mathParaPr = para.Descendants(M + "oMathParaPr").FirstOrDefault();
        var mathAlign = mathParaPr?.Element(M + "jc")?.Attribute(M + "val")?.Value;

        if (IsCentered(paragraphAlign) && (mathParaPr == null || IsCentered(mathAlign))) return;

        SetAlignment(pPr.Element(W + "jc"), W + "jc", pPr, "center");
        if (mathParaPr != null)
        {
            SetAlignment(mathParaPr.Element(M + "jc"), M + "jc", mathParaPr, "center", M + "val");
        }

        fixes.Add(new FormulaFix(paragraphIndex, formulaIndex, "F1",
            $"公式位置：{DescribeAlignment(paragraphAlign)} → 居中"));
    }

    /// <summary>F5：公式段段前段后 0 磅。行距不动——R11 对含公式的段落有豁免。</summary>
    private static void FixSpacing(XElement para, int paragraphIndex, int formulaIndex, List<FormulaFix> fixes)
    {
        var pPr = para.Element(W + "pPr");
        var spacing = pPr?.Element(W + "spacing");
        if (spacing == null) return;

        var before = ParseInt(spacing.Attribute(W + "before")?.Value);
        var after = ParseInt(spacing.Attribute(W + "after")?.Value);
        if (before == 0 && after == 0) return;

        spacing.SetAttributeValue(W + "before", "0");
        spacing.SetAttributeValue(W + "after", "0");

        // beforeLines / afterLines 优先于磅值，留着它们等于没改。
        spacing.Attribute(W + "beforeLines")?.Remove();
        spacing.Attribute(W + "afterLines")?.Remove();

        fixes.Add(new FormulaFix(paragraphIndex, formulaIndex, "F5",
            $"段前 {before / 20.0:0.#} 磅、段后 {after / 20.0:0.#} 磅 → 均 0 磅"));
    }

    /// <summary>F4：公式内字体回到 Cambria Math。只动显式写死的 ascii/hAnsi，eastAsia 不碰。</summary>
    private static void FixMathFont(XElement para, int paragraphIndex, int formulaIndex, List<FormulaFix> fixes)
    {
        var foreign = new List<string>();

        foreach (var run in para.Descendants(M + "r"))
        {
            var rFonts = run.Element(W + "rPr")?.Element(W + "rFonts");
            var ascii = rFonts?.Attribute(W + "ascii")?.Value;
            if (rFonts == null || string.IsNullOrEmpty(ascii)) continue;
            if (string.Equals(ascii, FormulaAnalyzer.CambriaMath, StringComparison.OrdinalIgnoreCase)) continue;

            if (!foreign.Contains(ascii, StringComparer.OrdinalIgnoreCase)) foreign.Add(ascii);
            rFonts.SetAttributeValue(W + "ascii", FormulaAnalyzer.CambriaMath);
            rFonts.SetAttributeValue(W + "hAnsi", FormulaAnalyzer.CambriaMath);
        }

        if (foreign.Count == 0) return;

        fixes.Add(new FormulaFix(paragraphIndex, formulaIndex, "F4",
            $"公式内字体 {string.Join("、", foreign)} → {FormulaAnalyzer.CambriaMath}"));
    }

    /// <summary>
    /// F2：式号右对齐。两步——把右制表位放到版心右缘；式号前面没有制表符就补一个。
    ///
    /// 补 <c>w:tab</c> 是在公式内部加元素，形式上超出了「只动 pPr/m:rPr」的范围。
    /// 之所以还是做：Word 自己写出来的合规公式就是这样
    /// （<c>&lt;m:r&gt;…&lt;w:tab/&gt;&lt;m:t&gt;(1)&lt;/m:t&gt;&lt;/m:r&gt;</c>），
    /// 而且不补的话「只加制表位」在版面上一点变化都没有，等于假装修好了。
    /// 补的位置严格限定在式号那个 m:t 之前，不碰任何 m:t 的文本。
    /// </summary>
    private static void FixNumberAlignment(
        XElement para,
        string? numberText,
        XElement? numberRun,
        bool numberInMath,
        int textWidth,
        int paragraphIndex,
        int formulaIndex,
        List<FormulaFix> fixes,
        List<string> skips)
    {
        if (numberText == null) return;

        var pPr = GetOrAddPPr(para);
        var tabs = pPr.Element(W + "tabs");
        var rightTab = tabs?.Elements(W + "tab")
            .FirstOrDefault(t => IsRight(t.Attribute(W + "val")?.Value));

        var currentPos = ParseInt(rightTab?.Attribute(W + "pos")?.Value);
        var tabStopOk = rightTab != null && currentPos >= textWidth - RightTabToleranceTwips;
        var tabCharOk = HasTabBeforeNumber(para, numberRun, numberText, numberInMath);

        if (tabStopOk && tabCharOk) return;

        // 整段右对齐的写法也能算「式号靠右」，检查器认，这里就不动它。
        var paragraphAlign = pPr.Element(W + "jc")?.Attribute(W + "val")?.Value;
        if (rightTab == null && IsRight(paragraphAlign)) return;

        if (numberRun == null)
        {
            skips.Add($"段落 {paragraphIndex}：识别到式号但定位不到所在的 run，跳过右对齐");
            return;
        }

        // 先把制表符补上再看制表位：两件事都不做的话，加个制表位在版面上毫无变化。
        if (!tabCharOk && !InsertTabBeforeNumber(numberRun, numberText, numberInMath))
        {
            skips.Add($"段落 {paragraphIndex}：式号所在 run 结构异常，制表符没敢插，需人工加一个 Tab");
            return;
        }

        if (!tabStopOk)
        {
            if (tabs == null)
            {
                tabs = new XElement(W + "tabs");
                pPr.Add(tabs);
            }

            if (rightTab == null)
            {
                tabs.Add(new XElement(W + "tab",
                    new XAttribute(W + "val", "right"),
                    new XAttribute(W + "pos", textWidth.ToString())));
            }
            else
            {
                rightTab.SetAttributeValue(W + "pos", textWidth.ToString());
            }
        }

        var actions = new List<string>();
        if (!tabStopOk)
        {
            actions.Add(rightTab == null
                ? $"补右制表位到版心右缘（{textWidth} twips）"
                : $"右制表位 {currentPos} → {textWidth} twips");
        }

        if (!tabCharOk) actions.Add("式号前补制表符");

        fixes.Add(new FormulaFix(paragraphIndex, formulaIndex, "F2", string.Join("，", actions)));
    }

    /// <summary>
    /// 式号前面是不是已经有制表符把它推向右缘。两种写法要分开看：
    /// 式号在公式对象里时，<c>w:tab</c> 就在这个 m:r 内部；
    /// 式号写成正文 run 时，<c>w:tab</c> 在公式和式号之间的另一个 run 里。
    /// 只认「公式之后、式号之前」这一段，段落更前面的制表符不算——那会把公式本身推开。
    /// </summary>
    private static bool HasTabBeforeNumber(
        XElement para,
        XElement? numberRun,
        string numberText,
        bool numberInMath)
    {
        if (numberRun == null) return false;

        if (numberInMath)
        {
            var nodes = numberRun.Nodes().ToList();
            var textIndex = nodes.FindIndex(node =>
                node is XElement element &&
                element.Name == M + "t" &&
                element.Value.Contains(numberText, StringComparison.Ordinal));
            if (textIndex < 0) return false;

            return nodes.Take(textIndex).Any(node => node is XElement element && element.Name == W + "tab");
        }

        var lastMath = para.Elements().LastOrDefault(element =>
            element.Name == M + "oMathPara" || element.Name == M + "oMath");
        var scope = lastMath ?? para;

        return scope.ElementsAfterSelf()
            .TakeWhile(element => element != numberRun)
            .Any(element => element.Descendants(W + "tab").Any());
    }

    /// <summary>在式号文字之前插一个 <c>w:tab</c>，位置紧贴 rPr 之后（没有 rPr 就插最前面）。</summary>
    private static bool InsertTabBeforeNumber(XElement numberRun, string numberText, bool numberInMath)
    {
        if (!numberInMath)
        {
            // 式号写成正文 run 时，制表符得单独一个 run，塞进文字 run 里 Word 不认。
            var tabRun = new XElement(W + "r", new XElement(W + "tab"));
            numberRun.AddBeforeSelf(tabRun);
            return true;
        }

        var firstText = numberRun.Elements(M + "t")
            .FirstOrDefault(t => t.Value.Contains(numberText, StringComparison.Ordinal));
        if (firstText == null) return false;

        firstText.AddBeforeSelf(new XElement(W + "tab"));
        return true;
    }

    /// <summary>
    /// F3：式号字体与字号。只处理式号写在正文 run 里的情况——
    /// 式号抄在公式对象里时，字体字号由公式引擎给，动它等于改公式本体的排版。
    /// 圆括号不在这里修：那是改文字，不在本期范围。
    /// </summary>
    private static void FixNumberFormat(
        string? numberText,
        XElement? numberRun,
        bool numberInMath,
        int paragraphIndex,
        int formulaIndex,
        List<FormulaFix> fixes,
        List<string> skips)
    {
        if (numberText == null) return;

        var missingParens = !(numberText.TrimStart().StartsWith('(') || numberText.TrimStart().StartsWith('（'))
            || !(numberText.TrimEnd().EndsWith(')') || numberText.TrimEnd().EndsWith('）'));
        if (missingParens)
        {
            skips.Add($"段落 {paragraphIndex}：式号「{numberText}」缺圆括号，加括号要改文字，本档不自动改");
        }

        if (numberRun == null || numberInMath) return;

        var rPr = numberRun.Element(W + "rPr");
        if (rPr == null) return;

        var rFonts = rPr.Element(W + "rFonts");
        var ascii = rFonts?.Attribute(W + "ascii")?.Value;
        var fontChanged = false;
        if (rFonts != null && !string.IsNullOrEmpty(ascii) &&
            !string.Equals(ascii, FormulaAnalyzer.TimesNewRoman, StringComparison.OrdinalIgnoreCase))
        {
            rFonts.SetAttributeValue(W + "ascii", FormulaAnalyzer.TimesNewRoman);
            rFonts.SetAttributeValue(W + "hAnsi", FormulaAnalyzer.TimesNewRoman);
            fixes.Add(new FormulaFix(paragraphIndex, formulaIndex, "F3",
                $"式号西文字体 {ascii} → {FormulaAnalyzer.TimesNewRoman}"));
            fontChanged = true;
        }

        var sizeChanged = false;
        foreach (var name in new[] { "sz", "szCs" })
        {
            var element = rPr.Element(W + name);
            if (element == null) continue;

            var halfPoints = ParseInt(element.Attribute(W + "val")?.Value);
            if (halfPoints <= 0 || halfPoints == FormulaAnalyzer.EquationNumberHalfPoints) continue;

            element.SetAttributeValue(W + "val", FormulaAnalyzer.EquationNumberHalfPoints.ToString());
            sizeChanged = true;
            if (name == "sz")
            {
                fixes.Add(new FormulaFix(paragraphIndex, formulaIndex, "F3",
                    $"式号字号 {halfPoints / 2.0:0.#} 磅 → 10.5 磅"));
            }
        }

        if (!fontChanged && !sizeChanged && numberRun.Element(W + "rPr")?.Element(W + "rFonts") == null)
        {
            // 字体没显式写，就是继承样式；凭空补一个 Times New Roman 会盖掉样式，
            // 而样式层的问题归 #117/#118，不在这里抢。
            skips.Add($"段落 {paragraphIndex}：式号字体走样式继承，本档不覆盖");
        }
    }

    private static XElement GetOrAddPPr(XElement para)
    {
        var pPr = para.Element(W + "pPr");
        if (pPr == null)
        {
            pPr = new XElement(W + "pPr");
            para.AddFirst(pPr);
        }

        return pPr;
    }

    private static void SetAlignment(
        XElement? existing,
        XName name,
        XElement parent,
        string value,
        XName? valueAttribute = null)
    {
        valueAttribute ??= W + "val";

        if (existing == null)
        {
            parent.Add(new XElement(name, new XAttribute(valueAttribute, value)));
            return;
        }

        existing.SetAttributeValue(valueAttribute, value);
    }

    private static XDocument ReadDocumentXml(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");

        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }

    private static void WriteDocumentXml(string inputPath, string outputPath, XDocument document)
    {
        if (!string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(inputPath, outputPath, overwrite: true);
        }

        using var stream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update);

        archive.GetEntry("word/document.xml")?.Delete();

        var newEntry = archive.CreateEntry("word/document.xml");
        using var writer = new StreamWriter(newEntry.Open());

        // 不要缩进：XDocument 默认会把整份文档重排，正文里平白多出一堆空白文本节点，
        // 既让 diff 没法看，也可能被别的工具当成内容。压缩输出更接近原文件。
        document.Save(writer, SaveOptions.DisableFormatting);
    }

    private static bool IsCentered(string? value) =>
        string.Equals(value, "center", StringComparison.OrdinalIgnoreCase);

    private static bool IsRight(string? value) =>
        string.Equals(value, "right", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "end", StringComparison.OrdinalIgnoreCase);

    private static string DescribeAlignment(string? value) => value switch
    {
        null or "" => "未设置对齐（默认左对齐）",
        "center" => "居中",
        "right" or "end" => "右对齐",
        "both" or "justify" => "两端对齐",
        "left" or "start" => "左对齐",
        _ => value,
    };

    private static int ParseInt(string? value) =>
        int.TryParse(value, out var parsed) ? parsed : 0;
}

public sealed record FormulaFormatResult(
    int StandaloneCount,
    int InlineCount,
    int TableRowCount,
    IReadOnlyList<FormulaFix> Fixes,
    IReadOnlyList<string> Skips);

/// <summary>一处实际写入的修改。<c>Code</c> 与检查器的 F1–F5 对齐，方便对账。</summary>
public sealed record FormulaFix(
    int ParagraphIndex,
    int FormulaIndex,
    string Code,
    string Description);
