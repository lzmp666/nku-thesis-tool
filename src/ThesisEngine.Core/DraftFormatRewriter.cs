using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Core;

/// <summary>
/// draft 规则的保守自动修正器。
///
/// 设计边界（严格遵守仓库口径）：
/// - 只消费各分析器「已经明确判定」的 Issue，不新增任何判据。分析器把判不出的维度
///   放进 NotChecked，不进 Issues，因此这些维度天然不会被本修复器触碰。
/// - 只改 Issue 指向段落的对应维度属性；Issue 的 Expected 是规范文本（如「宋体」「10.5磅」
///   「居中」「固定值20磅」「段前6磅」「2 个汉字符」），这里把文本映射回 OOXML 值。
///   映射不到的维度直接跳过并计入 Skipped，不猜、不硬判。
/// - 写段落级直接格式（pPr/rPr），与分析器 PartFormatReader 的读取口径一致：显式值优先生效，
///   修复后重新 Analyze 应读到规范值。不改样式表、不改文字、不改部件结构。
/// - 幂等：目标值与现值一致时不产生改动；重复运行输出字节一致。
/// - 规则状态 draft，本修复器只做「机械地把 Actual 对齐到 Expected」，不构成规范合规结论。
/// </summary>
public sealed class DraftFormatRewriter
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>一次修正的结果统计。</summary>
    public sealed record FixResult(
        int Applied,
        IReadOnlyList<string> Skipped,
        IReadOnlyList<string> ChangedParagraphs)
    {
        public bool HasChanges => Applied > 0;
    }

    /// <summary>
    /// 依据六个 draft 分析器的判定结果修正文档。
    /// 仅修改 word/document.xml 中被明确判定问题的段落；其余部件原样复制。
    /// </summary>
    public FixResult Apply(string docxPath, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(docxPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        if (!File.Exists(docxPath)) throw new FileNotFoundException("找不到文档", docxPath);

        var skipped = new List<string>();
        var changed = new List<string>();
        var applied = 0;

        // 跨分析器冲突检测：同一段落同一「属性槽」若被多个分析器给出不同判定，一律跳过，
        // 不带着冲突口径硬改。spacing（行距/段距）是正文与标题分析器都可能触及的槽，
        // 冲突时宁可不修，也不能把标题的合法行距改成正文的固定值。
        // 槽：font-east / font-west / size / bold / align / spacing / indent
        var claims = new Dictionary<(int Para, string Slot), (string Analyzer, string Expected)>();
        var conflicts = new HashSet<(int Para, string Slot)>();
        void Claim(int para, string dimension, string analyzer, string expected)
        {
            var slot = SlotOf(dimension);
            if (slot is null) return;
            var key = (para, slot);
            if (claims.TryGetValue(key, out var prior))
            {
                // 同一分析器内部对同一槽报多个子维度（如标题的行距+段前+段后）不算冲突；
                // 只有不同分析器口径相左时，才按冲突处理、整槽跳过。
                if (!string.Equals(prior.Analyzer, analyzer, StringComparison.Ordinal)
                    && !string.Equals(prior.Expected, expected, StringComparison.Ordinal))
                {
                    conflicts.Add(key);
                    skipped.Add($"段落 {para} 的 {slot} 被 {prior.Analyzer}（「{prior.Expected}」）与 {analyzer}（「{expected}」）给出不同期望，冲突跳过");
                }
            }
            else claims[key] = (analyzer, expected);
        }

        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var document = ReadPart(archive, "word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");

        var paragraphs = document.Root?.Descendants(W + "p").ToList()
            ?? new List<XElement>();

        // 先把六个分析器的全部 Issue 收集成统一列表（段落, 分析器, 码, 维度, 期望）。
        var pending = new List<(int Para, string Analyzer, string Code, string Dimension, string Expected)>();

        // 角色门控：被标题分析器识别为分级标题的段落，不能被正文分析器当正文修改。
        // 正文分析器靠文本锚点识别边界，编号标题（如「1.1 研究背景」）可能被它误纳入正文；
        // 标题分析器靠 outlineLvl/样式识别更可靠。两者冲突时以标题识别为准，跳过正文修复。
        var headingResult = new HeadingFormatAnalyzer().Analyze(docxPath);
        var headingParagraphs = headingResult.RecognizedHeadings.Select(h => h.ParagraphIndex).ToHashSet();

        // 1. 正文段落（R11）
        foreach (var i in new BodyTextFormatAnalyzer().Analyze(docxPath).Issues)
        {
            if (headingParagraphs.Contains(i.ParagraphIndex)) continue; // 已识别为标题，不当正文修
            var dimension = i.Code switch
            {
                "T1" => "Align", "T2" => "Indent", "T3" => "Line", "T4" => "Space", _
                    => ""
            };
            pending.Add((i.ParagraphIndex, "BodyTextFormatAnalyzer", i.Code, dimension, i.Expected));
        }
        // 2. 图题表题（R12）
        foreach (var i in new CaptionFormatAnalyzer().Analyze(docxPath).Issues)
            pending.Add((i.ParagraphIndex, "CaptionFormatAnalyzer", i.Code, i.Dimension, i.Expected));
        // 3. 标题层级（R7-R10）
        foreach (var i in headingResult.Issues)
            pending.Add((i.ParagraphIndex, "HeadingFormatAnalyzer", i.Code, i.Dimension, i.Expected));
        // 4. 目录（R6）
        foreach (var i in new TocFormatAnalyzer().Analyze(docxPath).Issues)
            pending.Add((i.ParagraphIndex, "TocFormatAnalyzer", i.Code, i.Dimension, i.Expected));
        // 5. 摘要（R4/R5）
        foreach (var i in new AbstractFormatAnalyzer().Analyze(docxPath).Issues)
            pending.Add((i.ParagraphIndex, "AbstractFormatAnalyzer", i.Code, AbstractDimension(i.Code, i.Title), i.Expected));
        // 6. 尾注部件（R14-R17）
        foreach (var i in new BackMatterFormatAnalyzer().Analyze(docxPath).Issues)
            pending.Add((i.ParagraphIndex, "BackMatterFormatAnalyzer", i.Code, BackMatterDimension(i.Title), i.Expected));

        // 第一遍：跨分析器冲突检测
        foreach (var item in pending)
            Claim(item.Para, item.Dimension, item.Analyzer, item.Expected);

        // 第二遍：应用无冲突的修正
        void ApplyIssue(int oneBasedIndex, string analyzer, string code, string dimension, string expected)
        {
            if (oneBasedIndex <= 0 || oneBasedIndex > paragraphs.Count)
            {
                skipped.Add($"{analyzer}/{code}：段落序号 {oneBasedIndex} 越界，跳过");
                return;
            }
            var slot = SlotOf(dimension);
            if (slot is not null && conflicts.Contains((oneBasedIndex, slot)))
            {
                skipped.Add($"{analyzer}/{code}：段落 {oneBasedIndex} 的 {slot} 存在跨分析器冲突，跳过");
                return;
            }
            var paragraph = paragraphs[oneBasedIndex - 1];
            if (TryApplyDimension(paragraph, dimension, expected))
            {
                applied++;
                if (!changed.Contains($"第 {oneBasedIndex} 段")) changed.Add($"第 {oneBasedIndex} 段");
            }
            else
            {
                skipped.Add($"{analyzer}/{code}：维度「{dimension}」期望「{expected}」无法安全映射，跳过");
            }
        }

        foreach (var item in pending)
            ApplyIssue(item.Para, item.Analyzer, item.Code, item.Dimension, item.Expected);

        if (applied > 0)
            WriteDocumentXml(docxPath, outputPath, document);
        else if (!string.Equals(Path.GetFullPath(docxPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
            CopyPackage(docxPath, outputPath);

        return new FixResult(applied, skipped.Distinct(StringComparer.Ordinal).ToArray(), changed);
    }

    /// <summary>把维度映射到它实际写入的 OOXML 属性槽，用于跨分析器冲突检测。</summary>
    private static string? SlotOf(string dimension) => dimension switch
    {
        "Font" or "FontEastAsia" => "font-east",
        "FontWestern" => "font-west",
        "Size" => "size",
        "Bold" => "bold",
        "Align" => "align",
        "Line" or "Before" or "After" or "Space" or "SpaceAndLine" => "spacing",
        "Indent" => "indent",
        _ => null
    };

    /// <summary>摘要分析器问题码 → 修正维度。码规则见 AbstractFormatAnalyzer：A1-A10，中英各 5 组。</summary>
    private static string AbstractDimension(string code, string title)
    {
        if (title.Contains("加粗", StringComparison.Ordinal)) return "Bold";
        if (title.Contains("对齐", StringComparison.Ordinal)) return "Align";
        if (title.Contains("行距", StringComparison.Ordinal) || title.Contains("段距", StringComparison.Ordinal)) return "SpaceAndLine";
        // 字体：英文摘要用西文槽（Arial/Times New Roman），中文摘要用东亚槽（宋体/黑体）。
        if (title.Contains("字体", StringComparison.Ordinal))
            return title.Contains("英文", StringComparison.Ordinal) ? "FontWestern" : "Font";
        if (title.Contains("字号", StringComparison.Ordinal)) return "Size";
        return "";
    }

    /// <summary>尾注部件问题码 → 修正维度，按标题关键字判定。</summary>
    private static string BackMatterDimension(string title)
    {
        if (title.Contains("中文字体", StringComparison.Ordinal)) return "FontEastAsia";
        if (title.Contains("西文字体", StringComparison.Ordinal)) return "FontWestern";
        if (title.Contains("字号", StringComparison.Ordinal)) return "Size";
        if (title.Contains("对齐", StringComparison.Ordinal)) return "Align";
        if (title.Contains("缩进", StringComparison.Ordinal)) return "Indent";
        if (title.Contains("段距", StringComparison.Ordinal) || title.Contains("行距", StringComparison.Ordinal)) return "SpaceAndLine";
        return "";
    }

    /// <summary>
    /// 把一个维度按期望文本写到段落上。返回 false 表示无法安全映射，调用方计入跳过。
    /// 只写显式值，不动样式表。
    /// </summary>
    private static bool TryApplyDimension(XElement paragraph, string dimension, string expected)
    {
        if (string.IsNullOrWhiteSpace(dimension)) return false;
        return dimension switch
        {
            "Font" => SetFont(paragraph, expected, western: false),
            "FontEastAsia" => SetFont(paragraph, expected, western: false),
            "FontWestern" => SetFont(paragraph, expected, western: true),
            "Size" => SetSize(paragraph, expected),
            "Bold" => SetBold(paragraph, expected),
            "Align" => SetAlign(paragraph, expected),
            "Line" => SetLine(paragraph, expected),
            "Before" => SetSingleSpace(paragraph, expected, before: true),
            "After" => SetSingleSpace(paragraph, expected, before: false),
            "Space" => SetZeroSpace(paragraph, expected),
            "SpaceAndLine" => SetSpaceAndLine(paragraph, expected),
            "Indent" => SetTwoCharIndent(paragraph, expected),
            _ => false
        };
    }

    private static XElement EnsureChild(XElement parent, XName name, string[]? order = null)
    {
        var existing = parent.Element(name);
        if (existing is not null) return existing;
        var created = new XElement(name);
        if (order is null) parent.Add(created);
        else InsertOrdered(parent, created, order);
        return created;
    }

    private static XElement PPr(XElement paragraph) => EnsureChild(paragraph, W + "pPr", ParagraphPropertySequence);

    /// <summary>
    /// 字符级属性（字体/字号/加粗）必须写到每个有文字的 run 的 rPr 上，
    /// 与 PartFormatReader.ResolveRun 的读取口径一致——它只读 run.Element(W+"rPr")，
    /// 不读 pPr 里的段落标记 rPr。写错位置会导致修复后分析器仍旧报问题。
    /// </summary>
    private static List<XElement> TextRunRPrs(XElement paragraph)
    {
        var result = new List<XElement>();
        foreach (var run in paragraph.Elements(W + "r"))
        {
            if (!run.Elements(W + "t").Any(t => !string.IsNullOrWhiteSpace(t.Value))) continue;
            var rPr = run.Element(W + "rPr");
            if (rPr is null)
            {
                rPr = new XElement(W + "rPr");
                run.AddFirst(rPr);
            }
            result.Add(rPr);
        }
        return result;
    }

    private static bool SetFont(XElement paragraph, string expected, bool western)
    {
        var font = NormalizeFont(expected);
        if (font is null) return false;
        var rPrs = TextRunRPrs(paragraph);
        if (rPrs.Count == 0) return false;
        foreach (var rPr in rPrs)
        {
            var rFonts = EnsureChild(rPr, W + "rFonts", RunPropertySequence);
            rFonts.SetAttributeValue(western ? W + "ascii" : W + "eastAsia", font);
            if (western) rFonts.SetAttributeValue(W + "hAnsi", font);
        }
        return true;
    }

    private static string? NormalizeFont(string expected)
    {
        var text = expected.Trim();
        if (PartFormatReader.IsSimSun(text)) return "宋体";
        if (PartFormatReader.IsSimHei(text)) return "黑体";
        if (PartFormatReader.IsArial(text)) return "Arial";
        if (PartFormatReader.IsTimesNewRoman(text)) return "Times New Roman";
        if (text.Contains("仿宋", StringComparison.Ordinal)
            || text.Equals("FangSong", StringComparison.OrdinalIgnoreCase)) return "仿宋";
        return null;
    }

    private static bool SetSize(XElement paragraph, string expected)
    {
        var half = ParseHalfPoints(expected);
        if (half is null) return false;
        var rPrs = TextRunRPrs(paragraph);
        if (rPrs.Count == 0) return false;
        foreach (var rPr in rPrs)
        {
            SetVal(rPr, "sz", half.Value);
            SetVal(rPr, "szCs", half.Value);
        }
        return true;
    }

    private static bool SetBold(XElement paragraph, string expected)
    {
        if (!expected.Contains("加粗", StringComparison.Ordinal)) return false;
        var rPrs = TextRunRPrs(paragraph);
        if (rPrs.Count == 0) return false;
        foreach (var rPr in rPrs)
        {
            var b = EnsureChild(rPr, W + "b", RunPropertySequence);
            b.SetAttributeValue(W + "val", "1");
        }
        return true;
    }

    private static bool SetAlign(XElement paragraph, string expected)
    {
        string? value = null;
        // 同时接受中文描述与 OOXML 原值（正文分析器 Expected 直接给 both）。
        var text = expected.Trim().ToLowerInvariant();
        if (text is "both" or "center" or "left" or "right" or "justify") value = text == "justify" ? "both" : text;
        else if (expected.Contains("居中", StringComparison.Ordinal)) value = "center";
        else if (expected.Contains("两端", StringComparison.Ordinal)) value = "both";
        else if (expected.Contains("左", StringComparison.Ordinal)) value = "left";
        else if (expected.Contains("右", StringComparison.Ordinal)) value = "right";
        if (value is null) return false;
        var pPr = PPr(paragraph);
        EnsureChild(pPr, W + "jc", ParagraphPropertySequence).SetAttributeValue(W + "val", value);
        return true;
    }

    /// <summary>行距维度：支持「单倍行距」与「固定值N磅」两种期望文本。</summary>
    private static bool SetLine(XElement paragraph, string expected)
    {
        var spacing = EnsureChild(PPr(paragraph), W + "spacing", ParagraphPropertySequence);
        if (expected.Contains("单倍行距", StringComparison.Ordinal))
        {
            spacing.SetAttributeValue(W + "lineRule", "auto");
            spacing.SetAttributeValue(W + "line", "240");
            return true;
        }
        var twips = ParseExactLineTwips(expected);
        if (twips is null) return false;
        spacing.SetAttributeValue(W + "lineRule", "exact");
        spacing.SetAttributeValue(W + "line", twips.Value.ToString(CultureInfo.InvariantCulture));
        return true;
    }

    /// <summary>单段距维度：期望形如「6磅」，before/after 由调用方指定。</summary>
    private static bool SetSingleSpace(XElement paragraph, string expected, bool before)
    {
        var match = Regex.Match(expected, @"([0-9]+(?:\.[0-9]+)?)\s*磅");
        if (!match.Success || !double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var pt))
            return false;
        var twips = (int)Math.Round(pt * 20, MidpointRounding.AwayFromZero);
        var spacing = EnsureChild(PPr(paragraph), W + "spacing", ParagraphPropertySequence);
        spacing.SetAttributeValue(before ? W + "before" : W + "after", twips.ToString(CultureInfo.InvariantCulture));
        return true;
    }

    private static bool SetZeroSpace(XElement paragraph, string expected)
    {
        // 正文 T4 的 Expected 是「段前0磅/段后0磅」，只处理段距，不碰行距。
        if (!expected.Contains("段前0磅", StringComparison.Ordinal) || !expected.Contains("段后0磅", StringComparison.Ordinal))
            return false;
        var spacing = EnsureChild(PPr(paragraph), W + "spacing", ParagraphPropertySequence);
        spacing.SetAttributeValue(W + "before", "0");
        spacing.SetAttributeValue(W + "after", "0");
        return true;
    }

    private static bool SetSpaceAndLine(XElement paragraph, string expected)
    {
        // Expected 形如「单倍行距，段前24磅/段后18磅」或「固定值20磅，段前0磅/段后0磅」。
        var spacing = EnsureChild(PPr(paragraph), W + "spacing", ParagraphPropertySequence);
        var changed = false;

        if (expected.Contains("单倍行距", StringComparison.Ordinal))
        {
            spacing.SetAttributeValue(W + "lineRule", "auto");
            spacing.SetAttributeValue(W + "line", "240");
            changed = true;
        }
        else if (ParseExactLineTwips(expected) is int exact)
        {
            spacing.SetAttributeValue(W + "lineRule", "exact");
            spacing.SetAttributeValue(W + "line", exact.ToString(CultureInfo.InvariantCulture));
            changed = true;
        }

        var before = ParseTaggedTwips(expected, "段前");
        var after = ParseTaggedTwips(expected, "段后");
        if (before is not null) { spacing.SetAttributeValue(W + "before", before.Value.ToString(CultureInfo.InvariantCulture)); changed = true; }
        if (after is not null) { spacing.SetAttributeValue(W + "after", after.Value.ToString(CultureInfo.InvariantCulture)); changed = true; }
        return changed;
    }

    private static bool SetTwoCharIndent(XElement paragraph, string expected)
    {
        // 正文首行缩进：Expected「2 个汉字符」→ firstLineChars=200，清掉 leftChars。
        // 目录左缩进：Expected「左缩进N汉字」→ leftChars=N*100，清掉 firstLineChars。
        var ind = EnsureChild(PPr(paragraph), W + "ind", ParagraphPropertySequence);
        var leftMatch = Regex.Match(expected, @"左缩进\s*([0-9]+)\s*汉字");
        if (leftMatch.Success && int.TryParse(leftMatch.Groups[1].Value, out var chars))
        {
            ind.SetAttributeValue(W + "leftChars", (chars * 100).ToString(CultureInfo.InvariantCulture));
            ind.SetAttributeValue(W + "left", (chars * 240).ToString(CultureInfo.InvariantCulture));
            ind.SetAttributeValue(W + "firstLineChars", "0");
            return true;
        }
        if (expected.Contains("2", StringComparison.Ordinal))
        {
            ind.SetAttributeValue(W + "firstLineChars", "200");
            ind.SetAttributeValue(W + "firstLine", "480");
            ind.SetAttributeValue(W + "leftChars", "0");
            return true;
        }
        return false;
    }

    private static void SetVal(XElement rPr, string localName, int value)
    {
        var element = EnsureChild(rPr, W + localName, RunPropertySequence);
        element.SetAttributeValue(W + "val", value.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>「10.5磅」→ 21（半点）。</summary>
    private static int? ParseHalfPoints(string text)
    {
        var match = Regex.Match(text, @"([0-9]+(?:\.[0-9]+)?)\s*磅");
        if (!match.Success) return null;
        if (!double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var pt)) return null;
        return (int)Math.Round(pt * 2, MidpointRounding.AwayFromZero);
    }

    /// <summary>「固定值20磅」→ 400 twips。</summary>
    private static int? ParseExactLineTwips(string text)
    {
        var match = Regex.Match(text, @"固定值\s*([0-9]+(?:\.[0-9]+)?)\s*磅");
        if (!match.Success) return null;
        if (!double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var pt)) return null;
        return (int)Math.Round(pt * 20, MidpointRounding.AwayFromZero);
    }

    /// <summary>从「段前6磅/段后12磅」中取指定标签的 twips。</summary>
    private static int? ParseTaggedTwips(string text, string tag)
    {
        var match = Regex.Match(text, tag + @"\s*([0-9]+(?:\.[0-9]+)?)\s*磅");
        if (!match.Success) return null;
        if (!double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var pt)) return null;
        return (int)Math.Round(pt * 20, MidpointRounding.AwayFromZero);
    }

    private static readonly string[] ParagraphPropertySequence =
    {
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
        "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap",
        "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr"
    };

    private static readonly string[] RunPropertySequence =
    {
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
        "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
        "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"
    };

    private static void InsertOrdered(XElement parent, XElement child, string[] order)
    {
        var childLocal = child.Name.LocalName;
        var childRank = Array.IndexOf(order, childLocal);
        if (childRank < 0) { parent.Add(child); return; }
        var next = parent.Elements().FirstOrDefault(e =>
        {
            var rank = Array.IndexOf(order, e.Name.LocalName);
            return rank >= 0 && rank > childRank;
        });
        if (next is null) parent.Add(child);
        else next.AddBeforeSelf(child);
    }

    private static XDocument? ReadPart(ZipArchive archive, string entry)
    {
        var item = archive.GetEntry(entry);
        if (item is null) return null;
        using var s = item.Open();
        return XDocument.Load(s);
    }

    private static void WriteDocumentXml(string inputPath, string outputPath, XDocument document)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        using var input = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var src = new ZipArchive(input, ZipArchiveMode.Read);
        using var output = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var dst = new ZipArchive(output, ZipArchiveMode.Create);
        foreach (var entry in src.Entries)
        {
            var target = dst.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            using var outStream = target.Open();
            if (entry.FullName == "word/document.xml")
            {
                using var writer = new StreamWriter(outStream, new System.Text.UTF8Encoding(false));
                document.Save(writer, SaveOptions.DisableFormatting);
            }
            else
            {
                using var inStream = entry.Open();
                inStream.CopyTo(outStream);
            }
        }
    }

    private static void CopyPackage(string inputPath, string outputPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        File.Copy(inputPath, outputPath, overwrite: true);
    }
}
