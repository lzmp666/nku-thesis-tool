using System.Text;
using System.Text.RegularExpressions;

namespace ThesisEngine.Core;

public enum ThesisSectionKind
{
    Unknown,
    FrontMatter,
    ChineseAbstract,
    EnglishAbstract,
    Contents,
    Body,
    References,
    Appendix,
    Acknowledgments
}

public enum ThesisBlockKind
{
    Unknown,
    Paragraph,
    Heading,
    Keywords,
    ContentsEntry,
    FigureCaption,
    TableCaption,
    TableCell
}

public sealed record ThesisStructureBlock(
    string BlockId,
    int ParagraphIndex,
    string Text,
    bool HasHeadingStyle = false,
    int? OutlineLevel = null,
    bool IsTableCell = false,
    CaptionKind? Caption = null,
    bool HasUnsupportedInlineContent = false);

/// <summary>
/// Confidence is an evidence-strength score, not a calibrated probability.
/// </summary>
public sealed record ThesisStructureDecision(
    string BlockId,
    int ParagraphIndex,
    ThesisSectionKind Section,
    ThesisBlockKind Kind,
    double Confidence,
    string Evidence,
    bool RequiresReview,
    int? HeadingLevel = null);

public sealed record ThesisStructureResult(
    IReadOnlyList<ThesisStructureDecision> Blocks,
    IReadOnlyList<string> Warnings);

/// <summary>
/// Conservative structural classifier for DOCX paragraphs, including documents whose
/// headings do not use Word heading styles. It only assigns a body region after finding
/// a bounded structural heading and never treats an otherwise unknown document as body.
/// </summary>
public sealed class ThesisStructureClassifier
{
    private static readonly Regex ChapterHeading = new(
        @"^第\s*(?<number>[0-9０-９零〇一二三四五六七八九十百千万两]+)\s*章\s*(?:[:：、.．\-—]\s*)?(?<title>.*)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex EnglishChapterHeading = new(
        @"^Chapter\s+(?<number>\d+(?:\.\d+)*|[IVXLCDM]+)(?:\s*[:：、.．\-—]?\s*(?<title>.*))?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex NumberedHeading = new(
        @"^(?<number>\d{1,3}(?:\.\d{1,3}){0,4})[.)、]?\s+(?<title>\S.{0,95})$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex UnnumberedBodyHeading = new(
        @"^(?:绪论|引言|研究背景|研究方法|方法|研究结果|结果|讨论|结论|结语|Introduction|Intro|Methods?|Results?|Discussion|Conclusions?)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex ContentsPageSuffix = new(
        @"^(?<core>.+?)(?:\s*(?:\.{2,}|…+|·{2,}|．{2,})\s*|\t+|\s{2,})\s*(?<page>\d{1,4})$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex ContentsSimplePageSuffix = new(
        @"^(?<core>.+?)\s+(?<page>\d{1,4})$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex ChineseKeywords = new(
        @"^(?:关键词|關鍵詞|关键字)\s*[:：]?\s*(?<value>.*)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex EnglishKeywords = new(
        @"^(?:key\s*words?|keywords?)\s*[:：]?\s*(?<value>.*)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex ChineseFigureCaption = new(
        @"^图\s*[A-ZＡ-Ｚ]?[0-9０-９]+(?:[.．-][0-9０-９]+)*(?:\s*[:：、.．]\s*|\s+)\S.{0,140}$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex ChineseTableCaption = new(
        @"^表\s*[A-ZＡ-Ｚ]?[0-9０-９]+(?:[.．-][0-9０-９]+)*(?:\s*[:：、.．]\s*|\s+)\S.{0,140}$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex EnglishFigureCaption = new(
        @"^(?:Figure|Fig\.?)\s*[0-9]+(?:[.\-][0-9]+)*(?:\s*[:：.\-]\s*|\s+)\S.{0,140}$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex EnglishTableCaption = new(
        @"^Table\s*[0-9]+(?:[.\-][0-9]+)*(?:\s*[:：.\-]\s*|\s+)\S.{0,140}$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex FrontMatterLabel = new(
        @"^(?:封面|独创性声明|原创性声明|学位论文使用授权书|使用授权书|版权声明|论文题目\s*[:：].*|作者\s*[:：].*|学号\s*[:：].*|学院(?:\s*[:：].*|$)|专业\s*[:：].*|学科门类\s*[:：].*|研究方向\s*[:：].*|指导教师\s*[:：].*|院系所\s*[:：].*|南开大学(?:研究生)?学位论文(?:\s*[（(].*[）)])?)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private const int MaximumHeadingLength = 100;
    private const int MaximumProseProbe = 5;

    public ThesisStructureResult Classify(IReadOnlyList<ThesisStructureBlock> blocks)
    {
        ArgumentNullException.ThrowIfNull(blocks);

        var warnings = new List<string>();
        if (blocks.Count == 0)
            return new ThesisStructureResult(Array.Empty<ThesisStructureDecision>(), new[] { "文档没有可分类的文本块。" });

        var ordered = blocks
            .Select((block, originalIndex) => new IndexedBlock(block, originalIndex))
            .OrderBy(item => item.Block.ParagraphIndex)
            .ThenBy(item => item.OriginalIndex)
            .ToArray();

        if (ordered.Select(item => item.Block.ParagraphIndex).Distinct().Count() != ordered.Length)
            warnings.Add("存在重复的段落索引；结果仍按输入顺序稳定处理重复项。");

        // Normalize and classify structural evidence once. Boundary probes then use
        // prefix/suffix indexes instead of rescanning the remainder for every heading.
        var analysis = new ClassificationAnalysis(ordered);

        var decisions = new List<ThesisStructureDecision>(ordered.Length);
        var section = ThesisSectionKind.Unknown;
        var inContents = false;
        var sawContentsEntry = false;
        var bodyFound = false;
        var bodyBoundaryConfidence = 0.0;
        var bodyRequiresReview = false;
        var unknownCount = 0;
        var unsupportedCount = 0;

        for (var i = 0; i < ordered.Length; i++)
        {
            var block = ordered[i].Block;
            var text = analysis.TextAt(i);

            if (block.IsTableCell)
            {
                decisions.Add(Create(block, section, ThesisBlockKind.TableCell, 0.99,
                    "输入结构明确标记为表格单元格；不按正文段落分类。", section == ThesisSectionKind.Unknown,
                    GetHeadingLevel(block, null)));
                continue;
            }

            if (block.Caption is { } explicitCaption)
            {
                var kind = explicitCaption == CaptionKind.Figure
                    ? ThesisBlockKind.FigureCaption
                    : ThesisBlockKind.TableCaption;
                decisions.Add(Create(block, section, kind, 0.99,
                    "输入结构明确提供了图/表题注类型；不按正文段落分类。", section == ThesisSectionKind.Unknown,
                    GetHeadingLevel(block, null)));
                continue;
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                unknownCount++;
                decisions.Add(Create(block, section, ThesisBlockKind.Unknown, 0.0,
                    "文本为空，无法判断章节或块类型。", true));
                continue;
            }

            if (block.HasUnsupportedInlineContent)
            {
                unsupportedCount++;
                unknownCount++;
                decisions.Add(Create(block, section, ThesisBlockKind.Unknown, 0.2,
                    "块含未支持的内联结构；不依据展平文本推断其章节或标题类型。", true));
                continue;
            }

            if (inContents && analysis.IsContentsEntry(i))
            {
                sawContentsEntry = true;
                decisions.Add(Create(block, ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry, 0.94,
                    "处于明确目录范围内，文本含标题式条目和页码/引导点。", false));
                continue;
            }

            var sectionTitle = analysis.TryGetSectionTitle(i, out var targetSection, out var normalizedTitle);
            if (inContents && sectionTitle)
            {
                // A standalone section name inside a TOC is commonly a TOC entry. Do not
                // let repeated labels such as References prematurely change the section.
                if (targetSection is ThesisSectionKind.ChineseAbstract or ThesisSectionKind.EnglishAbstract &&
                    analysis.HasSectionBodyContinuation(i))
                {
                    inContents = false;
                    sawContentsEntry = false;
                }
                else if (sawContentsEntry || targetSection is ThesisSectionKind.References or ThesisSectionKind.Appendix or ThesisSectionKind.Acknowledgments)
                {
                    decisions.Add(Create(block, ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry, 0.72,
                        "该段是目录中的独立章节名称，但没有可核验的页码引导信息；保留目录分类并要求复核。", true));
                    continue;
                }
                else
                {
                    inContents = false;
                    sawContentsEntry = false;
                }
            }

            if (sectionTitle)
            {
                section = targetSection;
                inContents = section == ThesisSectionKind.Contents;
                sawContentsEntry = false;
                bodyBoundaryConfidence = 0.0;
                bodyRequiresReview = false;
                var evidence = "独立章节标签与已知中/英文标题精确匹配：" + normalizedTitle + "。";
                decisions.Add(Create(block, section, ThesisBlockKind.Heading, 0.92, evidence,
                    section == ThesisSectionKind.Unknown, GetHeadingLevel(block, 1)));
                continue;
            }

            if (TryKeywords(text))
            {
                var keywordSection = section is ThesisSectionKind.ChineseAbstract or ThesisSectionKind.EnglishAbstract
                    ? section
                    : section == ThesisSectionKind.Unknown ? ThesisSectionKind.FrontMatter : section;
                var associated = section is ThesisSectionKind.ChineseAbstract or ThesisSectionKind.EnglishAbstract;
                if (keywordSection == ThesisSectionKind.FrontMatter && section == ThesisSectionKind.Unknown)
                    section = ThesisSectionKind.FrontMatter;
                decisions.Add(Create(block, keywordSection, ThesisBlockKind.Keywords,
                    associated ? 0.91 : 0.68,
                    associated
                        ? "关键词标签位于已识别的摘要范围内。"
                        : "关键词标签可识别，但没有可靠摘要标题确定其语言/归属。",
                    !associated));
                continue;
            }

            if (inContents)
            {
                if (analysis.TryGetHeading(i, out var contentsHeading) &&
                    TryGetBodyStartEvidence(contentsHeading, block, analysis, i,
                        restrictToContentsWindow: true, out var bodyStartEvidence))
                {
                    inContents = false;
                    sawContentsEntry = false;
                    section = ThesisSectionKind.Body;
                    bodyFound = true;
                    bodyBoundaryConfidence = bodyStartEvidence.Confidence;
                    bodyRequiresReview = bodyStartEvidence.RequiresReview;
                    decisions.Add(Create(block, section, ThesisBlockKind.Heading, bodyBoundaryConfidence,
                        bodyStartEvidence.Evidence,
                        bodyRequiresReview,
                        GetHeadingLevel(block, contentsHeading.Level)));
                    continue;
                }

                if (analysis.TryGetHeading(i, out _) || block.HasHeadingStyle || block.OutlineLevel.HasValue)
                {
                    sawContentsEntry = true;
                    decisions.Add(Create(block, ThesisSectionKind.Contents, ThesisBlockKind.ContentsEntry, 0.67,
                        "位于目录范围内且具有标题形态，但没有足够证据判断为真实正文起点。", true,
                        GetHeadingLevel(block, null)));
                    continue;
                }

                unknownCount++;
                decisions.Add(Create(block, ThesisSectionKind.Contents, ThesisBlockKind.Unknown, 0.4,
                    "位于目录范围内，但不符合可核验的目录条目形态。", true));
                continue;
            }

            if (analysis.TryGetTextCaption(i, out var textCaptionKind))
            {
                decisions.Add(Create(block, section,
                    textCaptionKind == CaptionKind.Figure ? ThesisBlockKind.FigureCaption : ThesisBlockKind.TableCaption,
                    0.74, "文本含图/表编号和题名形式，但没有题注样式或配对对象元数据；需要复核。", true));
                continue;
            }

            if (analysis.TryGetHeading(i, out var heading))
            {
                var hasStyleEvidence = block.HasHeadingStyle || block.OutlineLevel.HasValue;
                var beginsBody = section is ThesisSectionKind.Unknown or ThesisSectionKind.FrontMatter or
                    ThesisSectionKind.ChineseAbstract or ThesisSectionKind.EnglishAbstract;

                if (section == ThesisSectionKind.Body)
                {
                    decisions.Add(Create(block, section, ThesisBlockKind.Heading,
                        bodyBoundaryConfidence,
                        "位于已识别的正文范围内，标题形态与当前章节结构相符。",
                        bodyRequiresReview,
                        GetHeadingLevel(block, heading.Level)));
                    continue;
                }

                if (section is not (ThesisSectionKind.References or ThesisSectionKind.Appendix or ThesisSectionKind.Acknowledgments) &&
                    beginsBody && TryGetBodyStartEvidence(heading, block, analysis, i,
                        restrictToContentsWindow: false, out var bodyStartEvidence))
                {
                    section = ThesisSectionKind.Body;
                    bodyFound = true;
                    bodyBoundaryConfidence = bodyStartEvidence.Confidence;
                    bodyRequiresReview = bodyStartEvidence.RequiresReview;
                    decisions.Add(Create(block, section, ThesisBlockKind.Heading,
                        bodyBoundaryConfidence,
                        bodyStartEvidence.Evidence,
                        bodyRequiresReview,
                        GetHeadingLevel(block, heading.Level)));
                    continue;
                }

                if (heading.IsUnnumberedSemantic)
                {
                    decisions.Add(Create(block, section, ThesisBlockKind.Heading, 0.66,
                        "识别到无编号章节词，但缺少足够的正文衔接或后续章节结构；保留待复核，不据此建立正文范围。",
                        true, GetHeadingLevel(block, heading.Level)));
                    continue;
                }

                decisions.Add(Create(block, section, ThesisBlockKind.Heading, 0.66,
                    "识别到标题形态，但缺少足够的正文衔接或文档级闭合证据；保留待复核。",
                    true, GetHeadingLevel(block, heading.Level)));
                continue;

            }

            if (block.HasHeadingStyle || block.OutlineLevel.HasValue)
            {
                var knownSection = section != ThesisSectionKind.Unknown;
                decisions.Add(Create(block, section, ThesisBlockKind.Heading,
                    knownSection ? 0.9 : 0.66,
                    "Word 标题样式/大纲级别明确；无法从文本可靠推断具体章节类型。",
                    !knownSection, GetHeadingLevel(block, null)));
                continue;
            }

            if (FrontMatterLabel.IsMatch(text))
            {
                if (section == ThesisSectionKind.Unknown) section = ThesisSectionKind.FrontMatter;
                decisions.Add(Create(block, section, ThesisBlockKind.Paragraph, 0.9,
                    "与明确的封面、声明或学位论文信息标签匹配。", false));
                continue;
            }

            if (section == ThesisSectionKind.Unknown)
            {
                unknownCount++;
                decisions.Add(Create(block, ThesisSectionKind.Unknown, ThesisBlockKind.Unknown, 0.08,
                    "没有标题样式、独立章节标签或可信编号边界；保留未知，不默认归为正文。", true));
                continue;
            }

            var paragraphConfidence = section == ThesisSectionKind.FrontMatter
                ? 0.76
                : section == ThesisSectionKind.Body
                    ? Math.Max(0.82, bodyBoundaryConfidence)
                    : 0.85;
            decisions.Add(Create(block, section, ThesisBlockKind.Paragraph, paragraphConfidence,
                bodyRequiresReview
                    ? "沿用正文候选范围，但起始证据不足；该正文块与范围边界一并要求复核。"
                    : "沿用已确认的章节边界；该块不是标题、题注或表格单元格。",
                section == ThesisSectionKind.Body && bodyRequiresReview));
        }

        if (!bodyFound)
            warnings.Add("未识别到可靠的正文起始标题；没有把全文默认分类为正文。");
        if (unknownCount > 0)
            warnings.Add($"有 {unknownCount} 个文本块无法可靠分类，已保留未知或要求复核。");
        if (unsupportedCount > 0)
            warnings.Add($"有 {unsupportedCount} 个文本块含未支持的内联结构，未据此推断结构。");

        return new ThesisStructureResult(decisions, warnings.Distinct(StringComparer.Ordinal).ToArray());
    }

    private static ThesisStructureDecision Create(
        ThesisStructureBlock block,
        ThesisSectionKind section,
        ThesisBlockKind kind,
        double confidence,
        string evidence,
        bool requiresReview,
        int? headingLevel = null) =>
        new(block.BlockId, block.ParagraphIndex, section, kind, confidence, evidence, requiresReview, headingLevel);

    private static string Normalize(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var normalized = text.Normalize(NormalizationForm.FormKC)
            .Replace('\u00A0', ' ')
            .Replace('\u200B', ' ')
            .Replace('\u200C', ' ')
            .Replace('\u200D', ' ')
            .Replace('\uFEFF', ' ');
        return Regex.Replace(normalized, @"[ \t\r\n]+", " ").Trim();
    }

    private static bool TrySectionTitle(string text, out ThesisSectionKind section, out string normalizedTitle)
    {
        normalizedTitle = StripSectionNumbering(text).Trim().TrimEnd(':', '：');
        var compact = Regex.Replace(normalizedTitle, @"\s+", string.Empty);
        if (Regex.IsMatch(compact, @"^(?:中文)?摘要(?:\((?:中文|Chinese|Abstract)\))?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            section = ThesisSectionKind.ChineseAbstract;
            return true;
        }
        if (Regex.IsMatch(compact, @"^(?:英文摘要|(?:English)?Abstract(?:\((?:English|英文)\))?)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            section = ThesisSectionKind.EnglishAbstract;
            return true;
        }
        if (Regex.IsMatch(compact, @"^(?:目录|TableofContents|Contents|TOC)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            section = ThesisSectionKind.Contents;
            return true;
        }
        if (Regex.IsMatch(compact, @"^(?:主要)?参考文献$|^(?:References|Bibliography)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            section = ThesisSectionKind.References;
            return true;
        }
        if (Regex.IsMatch(compact, @"^附录(?:[A-Z0-9一二三四五六七八九十]+)?$|^Appendix(?:es)?(?:[A-Z0-9]+)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) ||
            Regex.IsMatch(normalizedTitle, @"^附录\s*[A-Z0-9一二三四五六七八九十]+(?:\s*[:：、.．\-—]?\s*\S.{0,75})?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) ||
            Regex.IsMatch(normalizedTitle, @"^Appendix(?:es)?\s+[A-Z0-9]+(?:\s*[:：、.．\-—]?\s*\S.{0,75})?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            section = ThesisSectionKind.Appendix;
            return true;
        }
        if (Regex.IsMatch(compact, @"^致谢$|^Acknowledg(?:e)?ments$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            section = ThesisSectionKind.Acknowledgments;
            return true;
        }

        section = ThesisSectionKind.Unknown;
        return false;
    }

    private static string StripSectionNumbering(string text)
    {
        var value = text.Trim();
        value = Regex.Replace(value,
            @"^(?:第\s*[0-9０-９零〇一二三四五六七八九十百千万两]+\s*章\s*|Chapter\s+(?:\d+(?:\.\d+)*|[IVXLCDM]+)\s*[:：、.．\-—]?\s*|\d{1,3}(?:\.\d{1,3}){0,4}[.)、]?\s+)",
            string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        return value;
    }

    private static bool TryKeywords(string text) => ChineseKeywords.IsMatch(text) || EnglishKeywords.IsMatch(text);

    private static bool TryContentsEntry(string text, out string core)
    {
        core = string.Empty;
        var match = ContentsPageSuffix.Match(text);
        if (!match.Success) match = ContentsSimplePageSuffix.Match(text);
        if (!match.Success) return false;

        var candidate = match.Groups["core"].Value.Trim();
        if (candidate.Length == 0 || candidate.Length > MaximumHeadingLength) return false;
        if (!IsHeadingCore(candidate) && !TrySectionTitle(candidate, out _, out _)) return false;
        core = candidate;
        return true;
    }

    private static bool IsHeadingCore(string text) => TryHeading(text, out _);

    private static bool TryTextCaption(string text, out CaptionKind kind)
    {
        if (ChineseFigureCaption.IsMatch(text) || EnglishFigureCaption.IsMatch(text))
        {
            kind = CaptionKind.Figure;
            return true;
        }
        if (ChineseTableCaption.IsMatch(text) || EnglishTableCaption.IsMatch(text))
        {
            kind = CaptionKind.Table;
            return true;
        }
        kind = default;
        return false;
    }

    private static bool TryHeading(string text, out HeadingShape shape)
    {
        shape = default;
        if (text.Length == 0 || text.Length > MaximumHeadingLength || ContainsSentencePunctuation(text)) return false;

        var cjk = ChapterHeading.Match(text);
        if (cjk.Success && IsPlausibleHeadingTail(cjk.Groups["title"].Value))
        {
            shape = new HeadingShape(1, true, Normalize(cjk.Groups["title"].Value), null);
            return true;
        }

        var english = EnglishChapterHeading.Match(text);
        if (english.Success && IsPlausibleHeadingTail(english.Groups["title"].Value))
        {
            var number = english.Groups["number"].Value;
            var depth = number.Count(character => character == '.') + 1;
            shape = new HeadingShape(depth, depth == 1, Normalize(english.Groups["title"].Value), number);
            return true;
        }

        var numbered = NumberedHeading.Match(text);
        if (numbered.Success && IsPlausibleHeadingTail(numbered.Groups["title"].Value))
        {
            var number = numbered.Groups["number"].Value;
            var level = number.Count(character => character == '.') + 1;
            shape = new HeadingShape(level, level == 1, Normalize(numbered.Groups["title"].Value), number);
            return true;
        }

        if (UnnumberedBodyHeading.IsMatch(text))
        {
            shape = new HeadingShape(1, false, text, null, true);
            return true;
        }

        return false;
    }

    private static bool IsPlausibleHeadingTail(string title)
    {
        if (title.Length == 0) return true;
        if (title.Length > 90 || ContainsSentencePunctuation(title)) return false;
        if (Regex.IsMatch(title, @"\b(?:according\s+to|we\s+(?:show|find|argue)|this\s+(?:study|paper|section))\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) return false;
        return true;
    }

    private static bool ContainsSentencePunctuation(string text) =>
        text.IndexOfAny(new[] { '。', '！', '？', '!', '?', ';', '；', '，', ',' }) >= 0;

    private static int? GetHeadingLevel(ThesisStructureBlock block, int? textLevel)
    {
        if (block.OutlineLevel is { } outline)
            return Math.Clamp(outline + 1, 1, 9);
        return textLevel is { } level ? Math.Clamp(level, 1, 9) : null;
    }

    private static bool TryGetBodyStartEvidence(
        HeadingShape heading,
        ThesisStructureBlock block,
        ClassificationAnalysis analysis,
        int index,
        bool restrictToContentsWindow,
        out BodyStartEvidence bodyStart)
    {
        bodyStart = default;
        var documentEvidence = analysis.AnalyzeBodyBoundary(index, restrictToContentsWindow);
        if (!documentEvidence.HasBodyProse) return false;

        var hasStyleEvidence = block.HasHeadingStyle || block.OutlineLevel.HasValue;
        if (heading.IsUnnumberedSemantic)
        {
            // A single word such as “Introduction” is not enough to claim the whole
            // following document as body. Require prose plus at least one more heading.
            if (!documentEvidence.HasAdditionalHeading) return false;

            var corroborated = documentEvidence.HasClosingSectionBoundary &&
                (documentEvidence.HasCoherentNumbering || documentEvidence.HasIntroductionHeading ||
                 documentEvidence.SemanticHeadingCount >= 3);
            bodyStart = new BodyStartEvidence(
                hasStyleEvidence ? 0.97 : corroborated ? 0.92 : 0.84,
                !hasStyleEvidence && !corroborated,
                corroborated
                    ? "无编号标题后有普通正文、连贯章节结构和独立文末章节边界，形成闭合的文档级正文证据。"
                    : hasStyleEvidence
                        ? "Word 标题样式与后续正文及其他章节结构共同确认无编号正文标题。"
                        : "无编号标题有后续正文和其他章节结构支撑，但缺少可信闭合边界；正文范围保留复核标记。");
            return true;
        }

        if (hasStyleEvidence && (heading.IsChapterRoot || documentEvidence.HasCoherentNumbering))
        {
            bodyStart = new BodyStartEvidence(0.97, false,
                "Word 标题样式/大纲级别与章节编号及后续普通正文相符。" +
                (documentEvidence.HasClosingSectionBoundary ? " 文档还出现独立文末章节边界。" : string.Empty));
            return true;
        }

        if (heading.IsChapterRoot)
        {
            var closed = documentEvidence.HasClosingSectionBoundary;
            bodyStart = new BodyStartEvidence(
                closed ? 0.94 : 0.84,
                !closed,
                closed
                    ? "无样式的第一章/Chapter 1 后有普通正文，并由独立参考文献、附录或致谢边界闭合。"
                    : "无样式的第一章/Chapter 1 后有普通正文，但没有独立参考文献、附录或致谢边界；正文范围要求复核。");
            return true;
        }

        if (documentEvidence.HasCoherentNumbering)
        {
            var closed = documentEvidence.HasClosingSectionBoundary;
            bodyStart = new BodyStartEvidence(
                closed ? 0.92 : 0.84,
                !closed,
                closed
                    ? "无样式编号标题形成连贯的多级编号序列，后有普通正文和独立文末章节边界。"
                    : "无样式编号标题形成连贯序列并有普通正文，但缺少独立文末章节边界；正文范围要求复核。");
            return true;
        }

        return false;
    }

    private static bool IsCoherentNumberingPair(string previousValue, string currentValue)
    {
        var previous = ParseNumberingPath(previousValue);
        var current = ParseNumberingPath(currentValue);
        if (previous.Length == 0 || current.Length == 0) return false;

        if (current.Length == previous.Length + 1 &&
            previous.SequenceEqual(current.Take(previous.Length)))
            return true;

        if (current.Length == previous.Length && current.Length > 1 &&
            previous.Take(previous.Length - 1).SequenceEqual(current.Take(current.Length - 1)) &&
            current[^1] == previous[^1] + 1)
            return true;

        return current.Length == 1 && previous.Length == 1 && current[0] == previous[0] + 1;
    }

    private static int[] ParseNumberingPath(string value) => value
        .Split('.', StringSplitOptions.RemoveEmptyEntries)
        .Select(component => int.TryParse(component, out var number) ? number : -1)
        .TakeWhile(number => number >= 0)
        .ToArray();

    private static bool IsIntroductionTitle(string title) =>
        Regex.IsMatch(title, @"^(?:绪论|引言|Introduction|Intro)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private readonly record struct BodyStartEvidence(double Confidence, bool RequiresReview, string Evidence);

    private readonly record struct BodyBoundaryEvidence(
        bool HasBodyProse,
        bool HasClosingSectionBoundary,
        bool HasAdditionalHeading,
        bool HasCoherentNumbering,
        bool HasIntroductionHeading,
        int SemanticHeadingCount);

    private static bool LooksLikeProse(string text) =>
        text.Length >= 35 || Regex.IsMatch(text, @"[。！？.!?；;，,]") ||
        Regex.Matches(text, @"\b[\p{L}]{2,}\b").Count >= 8;

    /// <summary>
    /// Per-call cache of normalized block evidence and O(1) range queries. No state is
    /// shared between classifier calls, so concurrent documents cannot contaminate one another.
    /// </summary>
    private sealed class ClassificationAnalysis
    {
        private readonly IndexedBlock[] _ordered;
        private readonly BlockEvidence[] _evidence;
        private readonly int[] _prosePrefix;
        private readonly int[] _headingPrefix;
        private readonly int[] _semanticHeadingPrefix;
        private readonly int[] _introductionPrefix;
        private readonly int[] _nextSectionBoundary;
        private readonly int[] _nextChapterRoot;
        private readonly int[] _nextCoherentPairCurrent;

        public ClassificationAnalysis(IndexedBlock[] ordered)
        {
            _ordered = ordered;
            _evidence = new BlockEvidence[ordered.Length];
            _prosePrefix = new int[ordered.Length + 1];
            _headingPrefix = new int[ordered.Length + 1];
            _semanticHeadingPrefix = new int[ordered.Length + 1];
            _introductionPrefix = new int[ordered.Length + 1];
            _nextSectionBoundary = new int[ordered.Length + 1];
            _nextChapterRoot = new int[ordered.Length + 1];
            _nextCoherentPairCurrent = new int[ordered.Length + 1];

            for (var index = 0; index < ordered.Length; index++)
            {
                var block = ordered[index].Block;
                var text = Normalize(block.Text);
                var eligible = !block.IsTableCell && !block.Caption.HasValue &&
                    !block.HasUnsupportedInlineContent && !string.IsNullOrWhiteSpace(text);
                var isContentsEntry = eligible && TryContentsEntry(text, out _);
                var section = ThesisSectionKind.Unknown;
                var normalizedTitle = string.Empty;
                var hasSectionTitle = eligible && !isContentsEntry &&
                    TrySectionTitle(text, out section, out normalizedTitle);
                var isTextCaption = eligible && !isContentsEntry && !hasSectionTitle && TryTextCaption(text, out _);
                HeadingShape? heading = eligible && !isContentsEntry && !hasSectionTitle && !isTextCaption &&
                    TryHeading(text, out var parsedHeading) ? parsedHeading : null;
                var isFrontMatter = eligible && !isContentsEntry && !hasSectionTitle && !isTextCaption &&
                    heading is null && FrontMatterLabel.IsMatch(text);
                var isProse = eligible && !isContentsEntry && !hasSectionTitle && !isTextCaption &&
                    heading is null && !isFrontMatter && LooksLikeProse(text);

                _evidence[index] = new BlockEvidence(text, eligible, isContentsEntry, hasSectionTitle,
                    section, normalizedTitle, isTextCaption, heading, isFrontMatter, isProse);
                _prosePrefix[index + 1] = _prosePrefix[index] + (isProse ? 1 : 0);
                _headingPrefix[index + 1] = _headingPrefix[index] + (heading.HasValue ? 1 : 0);
                var isSemantic = heading?.IsUnnumberedSemantic == true;
                _semanticHeadingPrefix[index + 1] = _semanticHeadingPrefix[index] + (isSemantic ? 1 : 0);
                _introductionPrefix[index + 1] = _introductionPrefix[index] +
                    (isSemantic && IsIntroductionTitle(heading!.Value.Title) ? 1 : 0);
            }

            BuildSectionBoundarySuffix();
            BuildChapterRootSuffix();
            BuildCoherentNumberingSuffix();
        }

        public string TextAt(int index) => _evidence[index].Text;
        public bool IsContentsEntry(int index) => _evidence[index].IsContentsEntry;
        public bool TryGetSectionTitle(int index, out ThesisSectionKind section, out string normalizedTitle)
        {
            var evidence = _evidence[index];
            section = evidence.Section;
            normalizedTitle = evidence.NormalizedTitle;
            return evidence.HasSectionTitle;
        }

        public bool TryGetTextCaption(int index, out CaptionKind kind)
        {
            var evidence = _evidence[index];
            if (evidence.IsTextCaption)
            {
                return TryTextCaption(evidence.Text, out kind);
            }
            kind = default;
            return false;
        }

        public bool TryGetHeading(int index, out HeadingShape heading)
        {
            if (_evidence[index].Heading is { } parsed)
            {
                heading = parsed;
                return true;
            }
            heading = default;
            return false;
        }

        public bool HasSectionBodyContinuation(int index)
        {
            var probed = 0;
            for (var nextIndex = index + 1; nextIndex < _ordered.Length && probed < MaximumProseProbe; nextIndex++)
            {
                var candidate = _ordered[nextIndex].Block;
                var evidence = _evidence[nextIndex];
                if (!evidence.IsEligible || evidence.IsContentsEntry) continue;
                if (evidence.HasSectionTitle) return false;
                if (candidate.HasHeadingStyle || candidate.OutlineLevel.HasValue || evidence.Heading.HasValue)
                {
                    probed++;
                    continue;
                }
                probed++;
                if (evidence.IsProse) return true;
            }
            return false;
        }

        public BodyBoundaryEvidence AnalyzeBodyBoundary(int index, bool restrictToContentsWindow)
        {
            var boundaryIndex = _nextSectionBoundary[index + 1];
            var end = boundaryIndex;
            var proseEnd = end;
            var closingBoundary = boundaryIndex < _ordered.Length &&
                IsClosingSection(_evidence[boundaryIndex].Section);

            if (restrictToContentsWindow)
            {
                // A heading shown in an unpaginated table of contents is not a body
                // start if another chapter root follows before real prose. Only inspect
                // five following blocks, and never borrow prose across that next root.
                var nextRoot = _nextChapterRoot[index + 1];
                if (nextRoot < proseEnd) proseEnd = nextRoot;
                var finiteEnd = Math.Min(_ordered.Length, index + 1 + MaximumProseProbe);
                if (finiteEnd < proseEnd) proseEnd = finiteEnd;
            }

            // Nearby prose proves this occurrence is the real chapter, while
            // a later references/appendix boundary independently closes its scope.
            var hasBodyProse = _prosePrefix[proseEnd] - _prosePrefix[index + 1] > 0;
            var hasAdditionalHeading = _headingPrefix[end] - _headingPrefix[index + 1] > 0;
            var semanticHeadingCount = _semanticHeadingPrefix[end] - _semanticHeadingPrefix[index];
            var hasIntroductionHeading = _introductionPrefix[end] - _introductionPrefix[index] > 0;
            var hasCoherentNumbering = _nextCoherentPairCurrent[index] < end;
            return new BodyBoundaryEvidence(
                hasBodyProse,
                hasBodyProse && closingBoundary,
                hasAdditionalHeading,
                hasCoherentNumbering,
                hasIntroductionHeading,
                semanticHeadingCount);
        }

        private void BuildSectionBoundarySuffix()
        {
            _nextSectionBoundary[^1] = _ordered.Length;
            for (var index = _ordered.Length - 1; index >= 0; index--)
                _nextSectionBoundary[index] = _evidence[index].HasSectionTitle
                    ? index
                    : _nextSectionBoundary[index + 1];
        }

        private void BuildChapterRootSuffix()
        {
            _nextChapterRoot[^1] = _ordered.Length;
            for (var index = _ordered.Length - 1; index >= 0; index--)
                _nextChapterRoot[index] = _evidence[index].Heading is { IsChapterRoot: true } &&
                    !_evidence[index].IsContentsEntry
                    ? index
                    : _nextChapterRoot[index + 1];
        }

        private void BuildCoherentNumberingSuffix()
        {
            var numberedHeadings = Enumerable.Range(0, _evidence.Length)
                .Where(index => !_evidence[index].IsContentsEntry && !_evidence[index].HasSectionTitle &&
                    _evidence[index].Heading?.Numbering is not null)
                .Select(index => (Index: index, Numbering: _evidence[index].Heading!.Value.Numbering!))
                .ToArray();
            var coherentPairByPreviousIndex = new Dictionary<int, int>();
            for (var i = 1; i < numberedHeadings.Length; i++)
            {
                var previous = numberedHeadings[i - 1];
                var current = numberedHeadings[i];
                if (IsCoherentNumberingPair(previous.Numbering, current.Numbering))
                    coherentPairByPreviousIndex[previous.Index] = current.Index;
            }

            _nextCoherentPairCurrent[^1] = _ordered.Length;
            for (var index = _ordered.Length - 1; index >= 0; index--)
            {
                var next = _nextCoherentPairCurrent[index + 1];
                if (coherentPairByPreviousIndex.TryGetValue(index, out var current))
                    next = Math.Min(next, current);
                _nextCoherentPairCurrent[index] = next;
            }
        }

        private static bool IsClosingSection(ThesisSectionKind section) => section is
            ThesisSectionKind.References or ThesisSectionKind.Appendix or ThesisSectionKind.Acknowledgments;

        private readonly record struct BlockEvidence(
            string Text,
            bool IsEligible,
            bool IsContentsEntry,
            bool HasSectionTitle,
            ThesisSectionKind Section,
            string NormalizedTitle,
            bool IsTextCaption,
            HeadingShape? Heading,
            bool IsFrontMatter,
            bool IsProse);
    }

    private readonly record struct IndexedBlock(ThesisStructureBlock Block, int OriginalIndex);
    private readonly record struct HeadingShape(
        int Level,
        bool IsChapterRoot,
        string Title,
        string? Numbering,
        bool IsUnnumberedSemantic = false);
}
