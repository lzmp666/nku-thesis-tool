using System.Globalization;
using System.Text.RegularExpressions;

namespace ThesisEngine.Core.Citation;

public sealed class CitationConsistencyAnalyzer
{
    private static readonly Regex CitationPattern = new(
        @"\[(?<numbers>\d+(?:\s*[,，、;；-]\s*\d+)*)\]",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex ReferencePattern = new(
        @"^\s*\[(?<number>\d+)\]\s*(?<body>.+?)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex JournalMarker = new(@"\[\s*J\s*\]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex YearPattern = new(@"(?<!\d)(?:1[5-9]|20\d{2})(?!\d)", RegexOptions.CultureInvariant);
    private static readonly Regex VolumeIssuePattern = new(@"\d+\s*[（(]\s*\d+\s*[）)]", RegexOptions.CultureInvariant);
    private static readonly Regex PagesPattern = new(
        @"(?:[:：]\s*\d+\s*[-–—]\s*\d+|\bpp?\.\s*\d+\s*[-–—]\s*\d+)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex FieldFingerprintPattern = new(@"[^\p{L}\p{N}]", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly HashSet<char> Punctuation = new("，。！？；：,.!?;:");

    public CitationAnalysisResult Analyze(string path) => Analyze(new CitationDocumentReader().Read(path));

    public CitationAnalysisResult Analyze(CitationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var findings = new List<CitationFinding>();
        var referenceHeading = FindReferencesHeading(document.Paragraphs);
        if (referenceHeading < 0)
        {
            findings.Add(new CitationFinding("citation/references-section", 0,
                "未识别到参考文献标题；无法核对正文引用与文末条目。", "缺少参考文献区"));
        }

        IReadOnlyList<ReferenceEntry> allReferences = referenceHeading < 0
            ? Array.Empty<ReferenceEntry>()
            : ReadReferences(document.Paragraphs, referenceHeading);
        CheckReferenceNumbers(allReferences, findings);
        var references = allReferences.Where(reference => reference.Number > 0).ToArray();
        var citations = referenceHeading < 0
            ? Array.Empty<CitationMarker>()
            : ReadCitations(document.Paragraphs, referenceHeading).ToArray();

        CheckSequentialNumbering(citations, findings);
        CheckOneToOneMapping(citations, references, findings);
        CheckCitationPosition(citations, findings);
        CheckReferenceFields(references, findings);
        CheckDuplicateWorks(references, findings);

        var unsupported = new List<string>();
        if (document.CitationFieldCount > 0)
            unsupported.Add("Zotero/EndNote 域内容仅统计数量，不解析域语义");
        if (document.Paragraphs.Count > 0)
            unsupported.Add("脚注和尾注不在本次检查范围");
        return new CitationAnalysisResult(findings, citations.Sum(citation => citation.Numbers.Count),
            allReferences.Count, document.CitationFieldCount, unsupported);
    }

    private static int FindReferencesHeading(IReadOnlyList<CitationParagraph> paragraphs)
    {
        for (var index = paragraphs.Count - 1; index >= 0; index--)
        {
            var title = NormalizeHeading(paragraphs[index].Text);
            if (title is "参考文献" or "主要参考文献" ||
                string.Equals(title, "references", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(title, "bibliography", StringComparison.OrdinalIgnoreCase))
                return index;
        }
        return -1;
    }

    private static string NormalizeHeading(string value) =>
        string.Concat(value.Where(character => !char.IsWhiteSpace(character) && character is not ':' and not '：'));

    private static List<ReferenceEntry> ReadReferences(IReadOnlyList<CitationParagraph> paragraphs, int headingIndex)
    {
        var entries = new List<ReferenceEntry>();
        for (var index = headingIndex + 1; index < paragraphs.Count; index++)
        {
            var paragraph = paragraphs[index];
            if (IsFollowingSection(paragraph.Text)) break;
            var match = ReferencePattern.Match(paragraph.Text);
            if (!match.Success || !int.TryParse(match.Groups["number"].Value, NumberStyles.None,
                    CultureInfo.InvariantCulture, out var number))
            {
                if (LooksLikeReferenceEntry(paragraph.Text))
                    entries.Add(new ReferenceEntry(0, paragraph.Text.Trim(), paragraph.ParagraphIndex));
                continue;
            }
            entries.Add(new ReferenceEntry(number, match.Groups["body"].Value, paragraph.ParagraphIndex));
        }
        return entries;
    }

    private static bool IsFollowingSection(string text)
    {
        var heading = NormalizeHeading(text);
        return heading.StartsWith("附录", StringComparison.Ordinal) ||
            heading is "致谢" or "勘误页" ||
            string.Equals(heading, "acknowledgements", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(heading, "appendix", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeReferenceEntry(string text) =>
        !string.IsNullOrWhiteSpace(text) &&
        (YearPattern.IsMatch(text) || JournalMarker.IsMatch(text) ||
         Regex.IsMatch(text, @"^\s*\d+[.、)]\s*\S", RegexOptions.CultureInvariant));

    private static IEnumerable<CitationMarker> ReadCitations(
        IReadOnlyList<CitationParagraph> paragraphs,
        int referenceHeading)
    {
        foreach (var paragraph in paragraphs.Take(referenceHeading))
        {
            foreach (Match match in CitationPattern.Matches(paragraph.Text))
            {
                var numbers = ParseNumbers(match.Groups["numbers"].Value);
                if (numbers.Count == 0) continue;
                var superscript = Enumerable.Range(match.Index, match.Length).All(position =>
                    paragraph.Runs.Any(run => run.IsSuperscript && position >= run.Start && position < run.Start + run.Length));
                if (!superscript && !LooksLikeCitation(paragraph.Text, match)) continue;
                yield return new CitationMarker(paragraph.ParagraphIndex, match.Index, match.Length,
                    match.Value, numbers, paragraph.Text);
            }
        }
    }

    private static bool LooksLikeCitation(string text, Match match)
    {
        var before = match.Index > 0 ? text[match.Index - 1] : '\0';
        var prefix = text[..match.Index];
        var suffix = text[(match.Index + match.Length)..];
        if (Regex.IsMatch(prefix, @"(?:数组\s*)?[A-Za-z_]\s*$", RegexOptions.CultureInvariant) &&
            Regex.IsMatch(suffix, @"^\s*(?:的)?\s*(?:下标|索引)", RegexOptions.CultureInvariant))
            return false;
        if (Regex.IsMatch(suffix,
                @"^\s*[（(][^）)]*(?:参数说明|数组下标|数组索引)[^）)]*[）)]\s*(?:表示|代表|指)",
                RegexOptions.CultureInvariant))
            return false;
        if (before is '_' || before is >= 'A' and <= 'Z' || before is >= 'a' and <= 'z' || char.IsDigit(before))
            return false;
        return true;
    }

    private static List<int> ParseNumbers(string value)
    {
        var pieces = Regex.Split(value, @"\s*[,，、;；]\s*");
        var result = new List<int>();
        foreach (var piece in pieces)
        {
            var range = piece.Split('-', StringSplitOptions.TrimEntries);
            if (range.Length == 1 && int.TryParse(range[0], out var number)) result.Add(number);
            else if (range.Length == 2 && int.TryParse(range[0], out var start) &&
                     int.TryParse(range[1], out var end) && end >= start && end - start <= 99)
                for (var current = start; current <= end; current++) result.Add(current);
        }
        return result.Distinct().ToList();
    }

    private static void CheckReferenceNumbers(IReadOnlyList<ReferenceEntry> references, List<CitationFinding> findings)
    {
        foreach (var duplicate in references.Where(reference => reference.Number > 0)
                     .GroupBy(reference => reference.Number).Where(group => group.Count() > 1))
        {
            foreach (var entry in duplicate.Skip(1))
                findings.Add(new CitationFinding("citation/reference-number-duplicate", entry.ParagraphIndex,
                    $"参考文献编号 [{duplicate.Key}] 重复。", $"[{duplicate.Key}]"));
        }
        foreach (var entry in references.Where(reference => reference.Number == 0))
            findings.Add(new CitationFinding("citation/reference-number-format", entry.ParagraphIndex,
                "疑似参考文献条目缺少顺序编码编号 [n]。", entry.Body));
    }

    private static void CheckSequentialNumbering(IReadOnlyList<CitationMarker> citations, List<CitationFinding> findings)
    {
        var seen = new HashSet<int>();
        var expected = 1;
        foreach (var citation in citations)
        {
            foreach (var number in citation.Numbers)
            {
                if (!seen.Add(number)) continue;
                if (number != expected)
                    findings.Add(new CitationFinding("citation/sequence", citation.ParagraphIndex,
                        $"正文引用首次出现的编号应连续；此处首次出现 [{number}]，预期 [{expected}]。", citation.Text));
                expected = number + 1;
            }
        }
    }

    private static void CheckOneToOneMapping(
        IReadOnlyList<CitationMarker> citations,
        IReadOnlyList<ReferenceEntry> references,
        List<CitationFinding> findings)
    {
        var cited = citations.SelectMany(citation => citation.Numbers).ToHashSet();
        var listed = references.Select(reference => reference.Number).ToHashSet();
        foreach (var number in cited.Except(listed).Order())
        {
            var citation = citations.First(item => item.Numbers.Contains(number));
            findings.Add(new CitationFinding("citation/missing-reference", citation.ParagraphIndex,
                $"正文引用 [{number}] 没有对应的文末参考文献条目。", $"[{number}]"));
        }
        foreach (var reference in references.Where(reference => !cited.Contains(reference.Number)))
            findings.Add(new CitationFinding("citation/uncited-reference", reference.ParagraphIndex,
                $"参考文献 [{reference.Number}] 未在正文中引用。", $"[{reference.Number}] {reference.Body}"));
    }

    private static void CheckCitationPosition(IReadOnlyList<CitationMarker> citations, List<CitationFinding> findings)
    {
        foreach (var citation in citations)
        {
            var previous = PreviousNonWhitespace(citation.ParagraphText, citation.Start - 1);
            if (previous >= 0 && Punctuation.Contains(citation.ParagraphText[previous]))
                findings.Add(new CitationFinding("citation/punctuation-position", citation.ParagraphIndex,
                    "引用标注位于标点之后；应检查是否应置于标点之前。", citation.Text));
        }
    }

    private static void CheckReferenceFields(IReadOnlyList<ReferenceEntry> references, List<CitationFinding> findings)
    {
        foreach (var reference in references)
        {
            var body = reference.Body;
            if (string.IsNullOrWhiteSpace(body))
                AddMissing("条目内容");
            if (!YearPattern.IsMatch(body)) AddMissing("年份");
            var authorLead = Regex.Split(body, @"\[(?:J|M|D|C|N|R|S|P|EB/OL|J/OL|M/OL)\]", RegexOptions.IgnoreCase)
                .FirstOrDefault()?.Trim() ?? body.Trim();
            if (!HasRecognizableAuthor(authorLead)) AddMissing("作者");
            if (JournalMarker.IsMatch(body))
            {
                if (!VolumeIssuePattern.IsMatch(body)) AddMissing("卷/期");
                if (!PagesPattern.IsMatch(body)) AddMissing("页码");
            }
            var authorSection = Regex.Split(body, @"\[(?:J|M|D|C|N|R|S|P|EB/OL|J/OL|M/OL)\]", RegexOptions.IgnoreCase)
                .FirstOrDefault() ?? body;
            var hasEtAl = Regex.IsMatch(authorSection, @"(?:等|et\s+al\.?)", RegexOptions.IgnoreCase);
            var chineseAuthorLead = Regex.Split(authorSection, @"[.。]").FirstOrDefault() ?? authorSection;
            var chineseAuthorCount = Regex.Split(chineseAuthorLead, @"[，、;；]").Count(part => !string.IsNullOrWhiteSpace(part));
            var englishAuthors = Regex.Matches(authorSection,
                @"(?:^|[,;]\s*)([A-Z][A-Za-z'’-]+)\s*,\s*[A-Z](?:[A-Z.]|\.)*(?=\s*[,;.])",
                RegexOptions.CultureInvariant);
            if ((chineseAuthorCount > 3 || englishAuthors.Count > 3) && !hasEtAl)
                AddMissing("超过三位作者时的「等」/“et al.”标记");

            void AddMissing(string field) => findings.Add(new CitationFinding("citation/reference-fields",
                reference.ParagraphIndex, $"参考文献 [{reference.Number}] 缺少可识别的{field}字段。",
                $"[{reference.Number}] {body}"));
        }
    }

    private static bool HasRecognizableAuthor(string authorLead)
    {
        var firstField = Regex.Split(authorLead, @"[.。]")[0].Trim();
        if (firstField.Length == 0) return false;
        if (Regex.IsMatch(firstField, @"^[\p{IsCJKUnifiedIdeographs}]{2,4}(?:\s*[，、;；]\s*[\p{IsCJKUnifiedIdeographs}]{2,4})*$",
                RegexOptions.CultureInvariant))
            return true;
        return Regex.IsMatch(firstField,
            @"^[A-Z][A-Za-z'’-]+(?:\s*,\s*[A-Z](?:[A-Z.]|\.)*)?(?:\s*[,;]\s*[A-Z][A-Za-z'’-]+(?:\s*,\s*[A-Z](?:[A-Z.]|\.)*)?)*$",
            RegexOptions.CultureInvariant);
    }

    private static void CheckDuplicateWorks(IReadOnlyList<ReferenceEntry> references, List<CitationFinding> findings)
    {
        var duplicateGroups = references
            .Select(reference => (Reference: reference, Fingerprint: Fingerprint(reference.Body)))
            .Where(item => item.Fingerprint.Length >= 12)
            .GroupBy(item => item.Fingerprint, StringComparer.Ordinal)
            .Where(group => group.Select(item => item.Reference.Number).Distinct().Count() > 1);
        foreach (var group in duplicateGroups)
        {
            foreach (var item in group.Skip(1))
                findings.Add(new CitationFinding("citation/duplicate-work", item.Reference.ParagraphIndex,
                    $"该参考文献内容与编号 [{group.First().Reference.Number}] 相同；同一文献应复用一个编号。",
                    $"[{item.Reference.Number}] {item.Reference.Body}"));
        }
    }

    private static string Fingerprint(string value) =>
        FieldFingerprintPattern.Replace(value, string.Empty).ToLowerInvariant();

    private static int PreviousNonWhitespace(string text, int index)
    {
        while (index >= 0 && char.IsWhiteSpace(text[index])) index--;
        return index;
    }

    private sealed record ReferenceEntry(int Number, string Body, int ParagraphIndex);
    private sealed record CitationMarker(int ParagraphIndex, int Start, int Length, string Text,
        IReadOnlyList<int> Numbers, string ParagraphText);
}
