using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ThesisEngine.Core.Formula;

/// <summary>
/// 公式排版分析器（#114 第一档：只查不改）。
///
/// 这里有一条必须守住的前提：公式要从 OMML 入口反向枚举，不能挂在段落 / run 遍历下面。
/// 纯 OMML 的公式段没有 <c>w:r</c>、也没有 <c>w:t</c>，任何「先取 run 文本再判断」的写法
/// 都会把整段跳过，于是公式全都不合规、报告却是 0 条。所以入口是
/// <c>m:oMathPara</c> / <c>m:oMath</c>，再由它向上找所在段落。
///
/// 检查项分两类，报告里也分开列，不混在一起：
/// - 有规范明文的：式号格式（R13）、式号右对齐（R13）、公式段段前段后（R11）。
/// - 排版惯例、规范无明文的：公式居中、公式内字体 Cambria Math。
///   规范里查不到这两条的条款号，就不给它编一个。
/// </summary>
public sealed class FormulaAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>式号：圆括号包数字，允许 1-1、(2.3) 这类章节编号。</summary>
    private static readonly Regex EquationNumberPattern =
        new(@"[（(]\s*\d+(?:\s*[-.．]\s*\d+)*\s*[)）]", RegexOptions.Compiled);

    /// <summary>
    /// 不带括号的式号。得单独认，否则「式号没加圆括号」这条永远报不出来——
    /// 连式号都识别不到，就只会当成「无式号」静静放过。
    /// </summary>
    private static readonly Regex BareNumberPattern =
        new(@"^\d+(?:\s*[-.．]\s*\d+)*$", RegexOptions.Compiled);

    /// <summary>A4 纵向宽度，twips。sectPr 缺失时的兜底。</summary>
    private const int DefaultPageWidthTwips = 11906;

    /// <summary>规范页边距左右各 3.2cm = 1814 twips，用于推算版心宽度。</summary>
    private const int DefaultMarginTwips = 1814;

    /// <summary>制表位与版心右缘的容差，500 twips ≈ 0.88cm。</summary>
    private const int TabStopToleranceTwips = 500;

    /// <summary>半磅单位下 10.5 磅 = 21。</summary>
    private const int EquationNumberHalfPoints = 21;

    private const string CambriaMath = "Cambria Math";
    private const string TimesNewRoman = "Times New Roman";

    public FormulaAnalysisResult Analyze(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var documentPart = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");

        XDocument doc;
        using (var reader = new StreamReader(documentPart.Open()))
        {
            doc = XDocument.Load(reader);
        }

        var body = doc.Root?.Element(W + "body")
            ?? throw new InvalidDataException("文档中没有 body");

        var textWidth = ReadTextWidthTwips(body);
        var paragraphs = body.Descendants(W + "p").ToList();

        var blocks = new List<FormulaBlock>();
        var issues = new List<FormulaIssue>();
        var inlineCount = 0;
        var formulaIndex = 0;

        for (var i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var paragraphIndex = i + 1;

            var topLevelMath = para.Descendants(M + "oMath")
                .Where(math => !math.Ancestors(M + "oMath").Any())
                .ToList();
            if (topLevelMath.Count == 0) continue;

            var mathText = string.Concat(topLevelMath.Select(ReadMathText));
            var layout = DetermineLayout(para, out var numberText, out var numberRun, out var numberInMath);
            if (numberInMath && numberText != null && mathText.EndsWith(numberText, StringComparison.Ordinal))
            {
                // 式号抄在公式对象里时它也出现在 m:t 序列末尾，预览里剥掉，避免和公式混在一起。
                mathText = mathText[..^numberText.Length];
            }

            if (layout == FormulaLayout.Inline)
            {
                // 行内公式不参与下面任何一项检查：它就是句子的一部分，
                // 居中、右对齐、独立成段这些要求对它没有意义。
                inlineCount++;
                continue;
            }

            formulaIndex++;
            var block = new FormulaBlock(
                Index: formulaIndex,
                ParagraphIndex: paragraphIndex,
                Text: Summarize(mathText),
                Layout: layout,
                HasEquationNumber: numberText != null,
                EquationNumber: numberText);
            blocks.Add(block);

            CheckCentering(para, block, layout, issues);
            CheckNumberAlignment(para, block, layout, numberText, textWidth, issues);
            CheckNumberFormat(block, numberRun, numberInMath, issues);
            CheckMathFont(para, block, issues);
            CheckSpacing(para, block, issues);
        }

        return new FormulaAnalysisResult(paragraphs.Count, blocks, issues, inlineCount);
    }

    /// <summary>
    /// 判断段落里是独立成段的公式还是行内公式，并顺带取出式号。
    /// 判据是「段落里不属于公式的正文文字还剩不剩」，不看有没有 oMathPara：
    /// Word 写 oMathPara、WPS 有时直接写 oMath，两种都得认。
    /// </summary>
    private static FormulaLayout DetermineLayout(
        XElement para,
        out string? numberText,
        out XElement? numberRun,
        out bool numberInMath)
    {
        numberText = null;
        numberRun = null;
        numberInMath = false;

        // 两列表格排版：公式一格、式号一格。表格排版下「段落居中」这条不适用，
        // 位置由列宽决定，所以单独作为一种布局，右对齐改看式号格的段落对齐。
        var table = para.Ancestors(W + "tbl").FirstOrDefault();
        if (table != null && table.Elements(W + "tr").Count() == 1)
        {
            var cells = table.Elements(W + "tr").First().Elements(W + "tc").ToList();
            if (cells.Count >= 2)
            {
                var lastCell = cells[^1];
                if (TryReadEquationNumber(ReadBodyText(lastCell), out var cellNumber))
                {
                    numberText = cellNumber;
                    numberRun = FindRunContaining(lastCell, cellNumber);
                    return FormulaLayout.TableRow;
                }
            }
        }

        // Word 的默认写法：式号抄在公式对象内部，形如
        // <m:r><w:rPr>…</w:rPr><w:tab/><m:t>(1)</m:t></m:r>
        // 这时式号是 m:t 不是 w:t，只读 w:t 的写法会认不出来，F2/F3 就永远不查。
        if (TryReadMathEquationNumber(para, out var mathNumber, out var mathRun))
        {
            numberText = mathNumber;
            numberRun = mathRun;
            numberInMath = true;
            return FormulaLayout.Standalone;
        }

        var bodyText = ReadBodyText(para);
        if (TryReadEquationNumber(bodyText, out var paragraphNumber))
        {
            numberText = paragraphNumber;
            numberRun = FindRunContaining(para, paragraphNumber);
            bodyText = bodyText.Replace(paragraphNumber, string.Empty, StringComparison.Ordinal);
        }

        // 去掉式号之后还有正文文字，说明公式只是句子的一部分。
        return string.IsNullOrWhiteSpace(bodyText)
            ? FormulaLayout.Standalone
            : FormulaLayout.Inline;
    }

    /// <summary>
    /// 在公式对象内部找式号。两种写法：
    /// 有右制表位时，最后一个含 <c>w:tab</c> 的 m:r 里制表符之后的文字就是式号；
    /// 没有制表位时，末段 m:t 若整段就是带圆括号的数字，也算式号。
    ///
    /// 第二种刻意只认带括号的：公式末尾经常是 <c>2</c>、<c>n</c> 这类单字符，
    /// 按裸数字认会把公式内容当式号，然后凭空报一条「式号未加圆括号」。
    /// </summary>
    private static bool TryReadMathEquationNumber(
        XElement para,
        out string? numberText,
        out XElement? numberRun)
    {
        numberText = null;
        numberRun = null;

        var mathRuns = para.Descendants(M + "r").ToList();
        if (mathRuns.Count == 0) return false;

        var tabRun = mathRuns.LastOrDefault(run => run.Elements(W + "tab").Any());
        if (tabRun != null)
        {
            var nodes = tabRun.Nodes().ToList();
            var tabIndex = nodes.FindIndex(node => node is XElement element && element.Name == W + "tab");
            if (tabIndex >= 0)
            {
                var afterTab = string.Concat(nodes
                    .Skip(tabIndex + 1)
                    .OfType<XElement>()
                    .Where(element => element.Name == M + "t")
                    .Select(element => element.Value));

                if (TryReadEquationNumber(afterTab, out var tabNumber))
                {
                    numberText = tabNumber;
                    numberRun = tabRun;
                    return true;
                }
            }
        }

        var lastRun = mathRuns[^1];
        var lastText = string.Concat(lastRun.Elements(M + "t").Select(t => t.Value)).Trim();
        var match = EquationNumberPattern.Match(lastText);
        if (!match.Success || match.Length != lastText.Length) return false;

        numberText = lastText;
        numberRun = lastRun;
        return true;
    }

    /// <summary>从一段文字里认出式号，带括号的和裸数字都认。</summary>
    private static bool TryReadEquationNumber(string text, out string numberText)
    {
        var match = EquationNumberPattern.Match(text);
        if (match.Success)
        {
            numberText = match.Value;
            return true;
        }

        var trimmed = text.Trim();
        if (BareNumberPattern.IsMatch(trimmed))
        {
            numberText = trimmed;
            return true;
        }

        numberText = string.Empty;
        return false;
    }

    /// <summary>公式居中（F1）。规范无明文条款，按排版惯例检查。</summary>
    private static void CheckCentering(
        XElement para,
        FormulaBlock block,
        FormulaLayout layout,
        List<FormulaIssue> issues)
    {
        // 表格排版不查居中：公式在单元格里，位置由列宽决定，报「未居中」只会误导。
        if (layout == FormulaLayout.TableRow) return;

        var paragraphAlign = para.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val")?.Value;
        var mathParaAlign = para.Descendants(M + "oMathParaPr")
            .Select(p => p.Element(M + "jc")?.Attribute(M + "val")?.Value)
            .FirstOrDefault(v => v != null);

        var centered = IsCentered(paragraphAlign) || IsCentered(mathParaAlign);
        if (centered) return;

        issues.Add(new FormulaIssue(
            Code: "F1",
            Title: "独立公式未居中",
            Basis: "排版惯例（规范 4.7 表4.6 未规定公式位置，本条不计入规范符合性）",
            HasClause: false,
            ParagraphIndex: block.ParagraphIndex,
            FormulaIndex: block.Index,
            Preview: block.Text,
            Current: DescribeAlignment(paragraphAlign ?? mathParaAlign),
            Expected: "居中"));
    }

    /// <summary>式号右对齐（F2）。R13 明文：序号右对齐。</summary>
    private static void CheckNumberAlignment(
        XElement para,
        FormulaBlock block,
        FormulaLayout layout,
        string? numberText,
        int textWidth,
        List<FormulaIssue> issues)
    {
        if (numberText == null) return;

        if (layout == FormulaLayout.TableRow)
        {
            // 式号在表格最后一格，看它的段落对齐。
            var table = para.Ancestors(W + "tbl").First();
            var lastCell = table.Elements(W + "tr").First().Elements(W + "tc").Last();
            var cellAlign = lastCell.Descendants(W + "jc")
                .Select(j => j.Attribute(W + "val")?.Value)
                .FirstOrDefault(v => v != null);
            if (IsRight(cellAlign)) return;

            issues.Add(new FormulaIssue(
                Code: "F2",
                Title: "式号未右对齐",
                Basis: "R13（4.7 表4.6）：序号右对齐",
                HasClause: true,
                ParagraphIndex: block.ParagraphIndex,
                FormulaIndex: block.Index,
                Preview: block.Text + " " + numberText,
                Current: cellAlign == null ? "未设置对齐（默认左对齐）" : DescribeAlignment(cellAlign),
                Expected: "右对齐"));
            return;
        }

        var pPr = para.Element(W + "pPr");
        var rightTab = pPr?.Element(W + "tabs")?
            .Elements(W + "tab")
            .FirstOrDefault(t => IsRight(t.Attribute(W + "val")?.Value));

        if (rightTab != null)
        {
            var pos = ParseInt(rightTab.Attribute(W + "pos")?.Value);
            if (pos >= textWidth - TabStopToleranceTwips) return;

            issues.Add(new FormulaIssue(
                Code: "F2",
                Title: "式号右对齐的制表位不在版心右缘",
                Basis: "R13（4.7 表4.6）：序号右对齐",
                HasClause: true,
                ParagraphIndex: block.ParagraphIndex,
                FormulaIndex: block.Index,
                Preview: block.Text + " " + numberText,
                Current: $"右制表位在第 {pos} twips（版心右缘 {textWidth} twips）",
                Expected: $"右制表位贴近版心右缘（约 {textWidth} twips）"));
            return;
        }

        // 没有制表位时，整段右对齐也算完成对齐（少数文档这么排）。
        var align = pPr?.Element(W + "jc")?.Attribute(W + "val")?.Value;
        if (IsRight(align)) return;

        issues.Add(new FormulaIssue(
            Code: "F2",
            Title: "式号未右对齐",
            Basis: "R13（4.7 表4.6）：序号右对齐",
            HasClause: true,
            ParagraphIndex: block.ParagraphIndex,
            FormulaIndex: block.Index,
            Preview: block.Text + " " + numberText,
            Current: "无右制表位，段落也未右对齐",
            Expected: "加右制表位到版心右缘，或整段右对齐"));
    }

    /// <summary>式号写法与字体（F3）。R13 明文：加圆括号、Times New Roman 10.5 磅。</summary>
    private static void CheckNumberFormat(
        FormulaBlock block,
        XElement? numberRun,
        bool numberInMath,
        List<FormulaIssue> issues)
    {
        if (block.EquationNumber == null) return;

        var inner = block.EquationNumber.Trim();
        var startsOk = inner.StartsWith('(') || inner.StartsWith('（');
        var endsOk = inner.EndsWith(')') || inner.EndsWith('）');
        if (!startsOk || !endsOk)
        {
            issues.Add(new FormulaIssue(
                Code: "F3",
                Title: "式号未加圆括号",
                Basis: "R13（4.7 表4.6）：序号加圆括号",
                HasClause: true,
                ParagraphIndex: block.ParagraphIndex,
                FormulaIndex: block.Index,
                Preview: block.Text,
                Current: block.EquationNumber,
                Expected: "(1) 形式"));
        }

        if (numberRun == null) return;

        // 只查显式设置。没显式写字体的 run 是继承样式，凭空断言「不是 Times New Roman」
        // 会造出一堆误报，而样式继承本身归样式检查管。
        var rPr = numberRun.Element(W + "rPr");

        // 式号抄在公式对象里时，它的字体由公式引擎给（实测 Word 写的就是 Cambria Math），
        // 不是用户的选择。照 R13 报「应为 Times New Roman」等于让所有用 Word 公式编辑器的人
        // 都背着一条改不掉的错，所以这种情况只查字号，不查字体。
        var ascii = rPr?.Element(W + "rFonts")?.Attribute(W + "ascii")?.Value;
        if (!numberInMath &&
            !string.IsNullOrEmpty(ascii) &&
            !string.Equals(ascii, TimesNewRoman, StringComparison.OrdinalIgnoreCase))
        {
            issues.Add(new FormulaIssue(
                Code: "F3",
                Title: "式号西文字体不对",
                Basis: "R13（4.7 表4.6）：Times New Roman 10.5 磅",
                HasClause: true,
                ParagraphIndex: block.ParagraphIndex,
                FormulaIndex: block.Index,
                Preview: block.EquationNumber!,
                Current: ascii,
                Expected: TimesNewRoman));
        }

        var halfPoints = ParseInt(rPr?.Element(W + "sz")?.Attribute(W + "val")?.Value);
        if (halfPoints > 0 && halfPoints != EquationNumberHalfPoints)
        {
            issues.Add(new FormulaIssue(
                Code: "F3",
                Title: "式号字号不对",
                Basis: "R13（4.7 表4.6）：Times New Roman 10.5 磅",
                HasClause: true,
                ParagraphIndex: block.ParagraphIndex,
                FormulaIndex: block.Index,
                Preview: block.EquationNumber!,
                Current: $"{halfPoints / 2.0:0.#} 磅",
                Expected: "10.5 磅"));
        }
    }

    /// <summary>公式内字体（F4）。规范无明文条款，按 Word 默认字体检查。</summary>
    private static void CheckMathFont(XElement para, FormulaBlock block, List<FormulaIssue> issues)
    {
        // WPS 存出来的公式经常没有 rPr，缺省即 Word 默认 Cambria Math，不报。
        var foreign = para.Descendants(M + "r")
            .Select(r => r.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii")?.Value)
            .Where(name => !string.IsNullOrEmpty(name))
            .Where(name => !string.Equals(name, CambriaMath, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (foreign.Count == 0) return;

        issues.Add(new FormulaIssue(
            Code: "F4",
            Title: "公式内字体不是 Cambria Math",
            Basis: "排版惯例（规范未规定公式字体，本条不计入规范符合性）",
            HasClause: false,
            ParagraphIndex: block.ParagraphIndex,
            FormulaIndex: block.Index,
            Preview: block.Text,
            Current: string.Join("、", foreign),
            Expected: CambriaMath));
    }

    /// <summary>公式段段前段后（F5）。R11 明文：段前 0 磅、段后 0 磅。</summary>
    private static void CheckSpacing(XElement para, FormulaBlock block, List<FormulaIssue> issues)
    {
        // 行距不查：R11 自带豁免——「段落中有数学表达式时，可根据表达需要设置该段的行距」。
        // 段前段后没有豁免，仍然按 0 磅查。
        var spacing = para.Element(W + "pPr")?.Element(W + "spacing");
        if (spacing == null) return;

        var before = ParseInt(spacing.Attribute(W + "before")?.Value);
        var after = ParseInt(spacing.Attribute(W + "after")?.Value);
        if (before == 0 && after == 0) return;

        issues.Add(new FormulaIssue(
            Code: "F5",
            Title: "公式段段前段后有间距",
            Basis: "R11（4.7 表4.6）：段前 0 磅、段后 0 磅",
            HasClause: true,
            ParagraphIndex: block.ParagraphIndex,
            FormulaIndex: block.Index,
            Preview: block.Text,
            Current: $"段前 {before / 20.0:0.#} 磅、段后 {after / 20.0:0.#} 磅",
            Expected: "段前 0 磅、段后 0 磅"));
    }

    /// <summary>版心宽度 = 页宽 - 左右页边距，单位 twips。</summary>
    private static int ReadTextWidthTwips(XElement body)
    {
        var sectPr = body.Elements(W + "sectPr").LastOrDefault()
            ?? body.Descendants(W + "sectPr").LastOrDefault();
        if (sectPr == null) return DefaultPageWidthTwips - DefaultMarginTwips * 2;

        var pageWidth = ParseInt(sectPr.Element(W + "pgSz")?.Attribute(W + "w")?.Value);
        if (pageWidth <= 0) pageWidth = DefaultPageWidthTwips;

        var pgMar = sectPr.Element(W + "pgMar");
        var left = ParseInt(pgMar?.Attribute(W + "left")?.Value);
        var right = ParseInt(pgMar?.Attribute(W + "right")?.Value);
        if (left <= 0) left = DefaultMarginTwips;
        if (right <= 0) right = DefaultMarginTwips;

        var width = pageWidth - left - right;
        return width > 0 ? width : DefaultPageWidthTwips - DefaultMarginTwips * 2;
    }

    private static string ReadMathText(XElement math) =>
        string.Concat(math.Descendants(M + "t").Select(t => t.Value));

    /// <summary>
    /// 取范围内「公式以外」的文字。公式里的文字是 m:t，本来就不在 w:t 里，
    /// 但 Word 偶尔会把公式的兜底文本也写进 w:t，所以再排除一次公式祖先。
    /// </summary>
    private static string ReadBodyText(XElement scope) =>
        string.Concat(scope.Descendants(W + "t")
            .Where(t => !t.Ancestors(M + "oMath").Any())
            .Select(t => t.Value));

    /// <summary>
    /// 找到包含某段文字的那个 run。w:r 和 m:r 都要找：
    /// 式号可能写成正文 run，也可能抄在公式对象的 m:r 里。
    /// </summary>
    private static XElement? FindRunContaining(XElement scope, string text)
    {
        return scope.Descendants()
            .Where(element => element.Name == W + "r" || element.Name == M + "r")
            .FirstOrDefault(run =>
            {
                var runText = string.Concat(
                    run.Elements(W + "t").Select(t => t.Value)
                        .Concat(run.Elements(M + "t").Select(t => t.Value)));
                return runText.Contains(text, StringComparison.Ordinal);
            });
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

    private static string Summarize(string text)
    {
        var compact = Regex.Replace(text, @"\s+", " ").Trim();
        if (compact.Length == 0) return "(无文本，纯符号公式)";
        return compact.Length <= 40 ? compact : compact[..40] + "…";
    }
}

/// <summary>公式在文档里的布局形态。</summary>
public enum FormulaLayout
{
    /// <summary>独占一段。</summary>
    Standalone,

    /// <summary>混在句子里的行内公式，不参与排版检查。</summary>
    Inline,

    /// <summary>单行两列（或多列）表格排版：公式一格、式号一格。</summary>
    TableRow,
}

public sealed record FormulaAnalysisResult(
    int ParagraphCount,
    IReadOnlyList<FormulaBlock> Blocks,
    IReadOnlyList<FormulaIssue> Issues,
    int InlineCount)
{
    public IEnumerable<FormulaIssue> ClauseIssues => Issues.Where(i => i.HasClause);

    public IEnumerable<FormulaIssue> ConventionIssues => Issues.Where(i => !i.HasClause);
}

public sealed record FormulaBlock(
    int Index,
    int ParagraphIndex,
    string Text,
    FormulaLayout Layout,
    bool HasEquationNumber,
    string? EquationNumber);

/// <summary>
/// 一条公式问题。<c>HasClause</c> 区分「规范明文」与「排版惯例」——
/// 惯例项不能算进规范符合性，报告里也单独列，避免拿惯例当校规去要求用户。
/// </summary>
public sealed record FormulaIssue(
    string Code,
    string Title,
    string Basis,
    bool HasClause,
    int ParagraphIndex,
    int? FormulaIndex,
    string Preview,
    string Current,
    string Expected);
