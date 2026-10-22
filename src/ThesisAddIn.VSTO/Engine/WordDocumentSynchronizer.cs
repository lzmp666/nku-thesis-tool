using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ThesisAddIn.Protocol;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>
    /// Applies engine patches and refreshes yellow source marks in the open Word document.
    /// No document is saved or reopened here, so Word keeps the edits in its normal undo stack
    /// and the author decides when to save the current document.
    /// </summary>
    public static class WordDocumentSynchronizer
    {
        private const string ToolAuthor = "NKU Thesis Tool";
        private const string ToolMarker = "[NKU-TOOL-ANNOTATION:v1]";
        private const string ToolBookmarkPrefix = "_NKU_MARK_";
        private const string WorkingCopyVariable = "NKUThesisToolWorkingCopy";
        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace W = WordNamespace;
        private static readonly Regex LocationPattern = new Regex(
            @"\Aword/document\.xml#paragraph\[(?<paragraph>\d+)\]/run\[(?<run>\d+)\]\z",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex OccurrencePattern = new Regex(
            @"\[NKU-OCCURRENCE:(?<location>word/document\.xml#paragraph\[\d+\]/run\[\d+\])\|(?<highlight>[a-zA-Z]+|absent)\]",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex SingleLocationPattern = new Regex(
            @"\[NKU-LOCATION:(?<location>word/document\.xml#paragraph\[\d+\]/run\[\d+\])\]",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex SingleHighlightPattern = new Regex(
            @"\[NKU-HIGHLIGHT:(?<highlight>[a-zA-Z]+|absent)\]",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private sealed class AnnotationRecord
        {
            public int Index { get; set; }
            public Dictionary<string, string> OriginalHighlights { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
        }

        private sealed class HighlightBookmarkRecord
        {
            public string Name { get; set; }
            public int OriginalHighlight { get; set; }
            public Word.Range Range { get; set; }
        }

        private sealed class XmlRun
        {
            public string Text { get; set; }
        }

        private sealed class XmlParagraph
        {
            public List<XmlRun> Runs { get; set; }
            public string Text { get; set; }
            public int CompletedTableRowsBefore { get; set; }
            public int TotalTableRows { get; set; }
        }

        public static bool IsWorkingCopy(Word.Document document)
        {
            foreach (Word.Variable variable in document.Variables)
                if (string.Equals(variable.Name, WorkingCopyVariable, StringComparison.Ordinal)
                    && string.Equals(variable.Value, "v1", StringComparison.Ordinal))
                    return true;
            return false;
        }

        public static void MarkWorkingCopy(Word.Document document)
        {
            if (IsWorkingCopy(document)) return;
            object version = "v1";
            document.Variables.Add(WorkingCopyVariable, ref version);
        }

        /// <summary>Refreshes diagnostic marks in the current document as one undoable Word action.</summary>
        public static bool RefreshAnnotations(Word.Document document, string snapshotPath,
            EngineAnalysisResult analysis, out string error)
        {
            return ExecuteUndoable(document, "刷新论文格式批注",
                () => ValidateAnnotationTargets(document, snapshotPath, analysis), markMutation =>
            {
                RefreshAnnotationsCore(document, snapshotPath, analysis, markMutation);
                document.Saved = false;
            }, out error);
        }

        /// <summary>Applies accepted patches and refreshes marks without replacing the open document.</summary>
        public static bool ApplyRepairsAndRefresh(Word.Document document, string snapshotPath,
            IReadOnlyList<EnginePatch> patches, EngineAnalysisResult analysis, out string error)
        {
            if (patches == null) throw new ArgumentNullException(nameof(patches));
            Dictionary<string, Word.Range> ranges = null;
            return ExecuteUndoable(document, "修复论文格式", () =>
            {
                ValidateAnnotationTargets(document, snapshotPath, analysis);
                ranges = BuildRanges(document, snapshotPath,
                    patches.Select(patch => patch.Location)
                        .Concat(GetDiagnosticLocations(analysis)));

                // Reject unsupported or stale work before changing the document.
                foreach (EnginePatch patch in patches)
                {
                    if (patch == null || !ranges.ContainsKey(patch.Location))
                        throw new InvalidDataException("修复计划中的位置无法映射到当前 Word 文本。");
                    ValidatePatch(ranges[patch.Location], patch);
                }

            }, markMutation =>
            {
                foreach (EnginePatch patch in patches)
                {
                    ApplyPatch(ranges[patch.Location], patch);
                    markMutation();
                }

                RefreshAnnotationsCore(document, snapshotPath, analysis, markMutation);
                document.Saved = false;
            }, out error);
        }

        public static bool TryGetRunRange(Word.Document document, string snapshotPath, string location,
            out Word.Range range, out string error)
        {
            range = null;
            error = null;
            try
            {
                range = BuildRanges(document, snapshotPath, new[] { location })[location];
                return true;
            }
            catch (Exception exception)
            {
                error = exception.Message;
                return false;
            }
        }

        public static Dictionary<string, Word.Range> GetOccurrenceRanges(Word.Document document,
            string snapshotPath, EngineAnalysisResult analysis)
        {
            return BuildRanges(document, snapshotPath, GetDiagnosticLocations(analysis));
        }

        private static bool ExecuteUndoable(Word.Document document, string actionName,
            Action validate, Action<Action> action, out string error)
        {
            error = null;
            if (document == null)
            {
                error = "Word 文档已关闭，请重新打开后重试。";
                return false;
            }

            dynamic undoRecord = null;
            bool customRecordStarted = false;
            bool mutated = false;
            try
            {
                // Validate before opening an undo record. An empty custom record is
                // absent from Word's undo stack: Undo would affect the preceding edit.
                validate();
                try
                {
                    undoRecord = document.Application.UndoRecord;
                    undoRecord.StartCustomRecord(actionName);
                    customRecordStarted = true;
                }
                catch (Exception exception)
                {
                    throw new InvalidOperationException("Word 无法创建撤销记录，已停止操作且未修改文档。", exception);
                }

                action(() => mutated = true);
                return true;
            }
            catch (Exception exception)
            {
                if (customRecordStarted)
                {
                    try
                    {
                        undoRecord.EndCustomRecord();
                        customRecordStarted = false;
                        if (mutated)
                        {
                            dynamic commandBars = document.Application.CommandBars;
                            commandBars.ExecuteMso("Undo");
                            error = "Word 未能完成操作，已撤销本次修改：" + exception.Message;
                        }
                        else
                        {
                            error = "Word 未能完成操作，文档内容未发生变化：" + exception.Message;
                        }
                    }
                    catch (Exception rollbackException)
                    {
                        error = "Word 未能完成操作：" + exception.Message + "；自动撤销也失败：" + rollbackException.Message;
                    }
                }
                else
                {
                    error = "Word 未能完成操作：" + exception.Message;
                }
                return false;
            }
            finally
            {
                if (customRecordStarted)
                {
                    try { undoRecord.EndCustomRecord(); }
                    catch { }
                }
            }
        }

        private static void ValidateAnnotationTargets(Word.Document document, string snapshotPath,
            EngineAnalysisResult analysis)
        {
            if (analysis == null) throw new ArgumentNullException(nameof(analysis));
            BuildRanges(document, snapshotPath, GetDiagnosticLocations(analysis)
                .Concat(ReadOwnedAnnotations(document).SelectMany(record => record.OriginalHighlights.Keys)));
        }

        private static void RefreshAnnotationsCore(Word.Document document, string snapshotPath,
            EngineAnalysisResult analysis, Action markMutation)
        {
            if (analysis == null) throw new ArgumentNullException(nameof(analysis));

            List<AnnotationRecord> existing = ReadOwnedAnnotations(document);
            List<HighlightBookmarkRecord> existingMarks = ReadOwnedHighlightBookmarks(document);
            var locations = new HashSet<string>(StringComparer.Ordinal);
            foreach (AnnotationRecord record in existing)
                foreach (string location in record.OriginalHighlights.Keys)
                    locations.Add(location);
            foreach (string location in GetDiagnosticLocations(analysis))
                locations.Add(location);

            Dictionary<string, Word.Range> ranges = BuildRanges(document, snapshotPath, locations);
            // Restore the highlight that was present before the tool marked each run.
            var originalByLocation = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (AnnotationRecord record in existing)
            {
                foreach (KeyValuePair<string, string> highlight in record.OriginalHighlights)
                {
                    if (originalByLocation.TryGetValue(highlight.Key, out string prior) && prior != highlight.Value)
                        throw new InvalidDataException("工具批注对同一位置记录了不同的原始高亮：" + highlight.Key);
                    originalByLocation[highlight.Key] = highlight.Value;
                }
            }

            foreach (KeyValuePair<string, string> highlight in originalByLocation)
            {
                if (string.Equals(GetHighlight(ranges[highlight.Key]), highlight.Value, StringComparison.Ordinal)) continue;
                SetHighlight(ranges[highlight.Key], highlight.Value);
                markMutation();
            }

            foreach (HighlightBookmarkRecord mark in existingMarks)
            {
                if (mark.Range == null) continue;
                if ((int)mark.Range.HighlightColorIndex != mark.OriginalHighlight)
                {
                    mark.Range.HighlightColorIndex = (Word.WdColorIndex)mark.OriginalHighlight;
                    markMutation();
                }
            }

            // Delete in reverse order so the 1-based collection indexes remain valid.
            foreach (int index in existing.Select(record => record.Index).OrderByDescending(index => index))
            {
                document.Comments[index].Delete();
                markMutation();
            }

            foreach (HighlightBookmarkRecord mark in existingMarks)
            {
                if (!document.Bookmarks.Exists(mark.Name)) continue;
                document.Bookmarks[mark.Name].Delete();
                markMutation();
            }

            // Comment removal can shift story ranges. Rebuild all targets against the same
            // text snapshot before writing new highlights and anchors.
            string[] newLocations = GetDiagnosticLocations(analysis).ToArray();
            ranges = BuildRanges(document, snapshotPath, newLocations);
            var freshOriginalHighlights = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (string location in newLocations.Distinct(StringComparer.Ordinal))
            {
                _ = GetHighlight(ranges[location]); // fail closed for mixed or unsupported highlights
                freshOriginalHighlights[location] = (int)ranges[location].HighlightColorIndex;
            }

            var markedLocations = new HashSet<string>(StringComparer.Ordinal);
            foreach (EngineDiagnostic diagnostic in analysis.Diagnostics ?? new List<EngineDiagnostic>())
            {
                string[] occurrences = GetOccurrences(diagnostic).ToArray();
                foreach (string location in occurrences)
                {
                    if (markedLocations.Add(location))
                        AddHighlightBookmark(document, location, ranges[location],
                            freshOriginalHighlights[location], markMutation);
                    if (ranges[location].HighlightColorIndex != Word.WdColorIndex.wdYellow)
                    {
                        ranges[location].HighlightColorIndex = Word.WdColorIndex.wdYellow;
                        markMutation();
                    }
                }
            }
            // This legacy marker identifies an analyzed document, not a separate copy.
            // Write it last; document variables do not create Word undo entries.
            MarkWorkingCopy(document);
        }

        private static List<AnnotationRecord> ReadOwnedAnnotations(Word.Document document)
        {
            var records = new List<AnnotationRecord>();
            for (int index = 1; index <= document.Comments.Count; index++)
            {
                Word.Comment comment = document.Comments[index];
                Word.Range range = comment.Range.Duplicate;
                range.TextRetrievalMode.IncludeHiddenText = true;
                string text = range.Text ?? string.Empty;
                if (text.IndexOf(ToolMarker, StringComparison.Ordinal) < 0)
                    continue;

                string author = comment.Author ?? string.Empty;
                if (!string.Equals(author, ToolAuthor, StringComparison.Ordinal))
                    throw new InvalidDataException("检测到带工具标记但作者不匹配的批注，为保护用户批注已停止刷新。");

                var record = new AnnotationRecord { Index = index };
                MatchCollection matches = OccurrencePattern.Matches(text);
                foreach (Match match in matches)
                    record.OriginalHighlights[match.Groups["location"].Value] = match.Groups["highlight"].Value;

                if (record.OriginalHighlights.Count == 0)
                {
                    Match location = SingleLocationPattern.Match(text);
                    Match highlight = SingleHighlightPattern.Match(text);
                    if (!location.Success || !highlight.Success)
                        throw new InvalidDataException("工具批注缺少位置或原始高亮元数据，为保护文档已停止刷新。");
                    record.OriginalHighlights[location.Groups["location"].Value] = highlight.Groups["highlight"].Value;
                }
                records.Add(record);
            }
            return records;
        }

        private static List<HighlightBookmarkRecord> ReadOwnedHighlightBookmarks(Word.Document document)
        {
            var records = new List<HighlightBookmarkRecord>();
            for (int index = document.Bookmarks.Count; index >= 1; index--)
            {
                Word.Bookmark bookmark = document.Bookmarks[index];
                string name = bookmark.Name ?? string.Empty;
                if (!name.StartsWith(ToolBookmarkPrefix, StringComparison.Ordinal)) continue;
                int separator = name.LastIndexOf('_');
                if (separator <= ToolBookmarkPrefix.Length ||
                    !int.TryParse(name.Substring(separator + 1), NumberStyles.HexNumber,
                        CultureInfo.InvariantCulture, out int originalHighlight))
                    throw new InvalidDataException("插件的高亮恢复标记无效；为避免改动文档，已停止刷新。");
                try { _ = GetHighlightName((Word.WdColorIndex)originalHighlight); }
                catch (InvalidDataException)
                {
                    throw new InvalidDataException("插件的高亮恢复颜色无效；为避免改动文档，已停止刷新。");
                }
                records.Add(new HighlightBookmarkRecord
                {
                    Name = name,
                    OriginalHighlight = originalHighlight,
                    Range = bookmark.Range.Duplicate
                });
            }
            return records;
        }

        private static void AddHighlightBookmark(Word.Document document, string location, Word.Range range,
            int originalHighlight, Action markMutation)
        {
            string locationHash;
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(location));
                locationHash = BitConverter.ToString(bytes).Replace("-", string.Empty).Substring(0, 16);
            }
            string name = ToolBookmarkPrefix + locationHash + "_" +
                originalHighlight.ToString("X8", CultureInfo.InvariantCulture);
            if (document.Bookmarks.Exists(name))
                throw new InvalidDataException("插件的高亮恢复标记已存在；请重新分析后再操作。");
            object bookmarkRange = range.Duplicate;
            document.Bookmarks.Add(name, ref bookmarkRange);
            markMutation();
        }

        private static Dictionary<string, Word.Range> BuildRanges(Word.Document document, string snapshotPath,
            IEnumerable<string> requestedLocations)
        {
            Dictionary<int, XmlParagraph> paragraphs = ReadXmlParagraphs(snapshotPath);
            int paragraphCount = document.Paragraphs.Count;
            int xmlParagraphCount = paragraphs.Count == 0 ? 0 : paragraphs.Keys.Max() + 1;
            int tableRowCount = paragraphs.Values.Select(paragraph => paragraph.TotalTableRows).DefaultIfEmpty(0).Max();
            if (paragraphCount != xmlParagraphCount + tableRowCount)
                throw new InvalidDataException("Word 段落与分析快照数量不一致（Word " + paragraphCount
                    + " 段，快照正文 " + xmlParagraphCount + " 段、表格 " + tableRowCount
                    + " 行），为避免改错位置已停止。");

            var result = new Dictionary<string, Word.Range>(StringComparer.Ordinal);
            var requestedByParagraph = new Dictionary<int, HashSet<int>>();
            foreach (string location in requestedLocations.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal))
            {
                Match match = LocationPattern.Match(location);
                if (!match.Success)
                    throw new InvalidDataException("暂不支持该位置格式：" + location);
                int paragraphIndex = int.Parse(match.Groups["paragraph"].Value, CultureInfo.InvariantCulture);
                int runIndex = int.Parse(match.Groups["run"].Value, CultureInfo.InvariantCulture);
                if (!paragraphs.TryGetValue(paragraphIndex, out XmlParagraph xmlParagraph) ||
                    runIndex < 0 || runIndex >= xmlParagraph.Runs.Count)
                    throw new InvalidDataException("分析位置已不存在：" + location);

                if (!requestedByParagraph.TryGetValue(paragraphIndex, out HashSet<int> requestedRuns))
                    requestedByParagraph[paragraphIndex] = requestedRuns = new HashSet<int>();
                requestedRuns.Add(runIndex);
            }

            foreach (KeyValuePair<int, HashSet<int>> requested in requestedByParagraph)
            {
                int paragraphIndex = requested.Key;
                XmlParagraph xmlParagraph = paragraphs[paragraphIndex];
                List<XmlRun> paragraphRuns = xmlParagraph.Runs;
                int wordParagraphIndex = paragraphIndex + xmlParagraph.CompletedTableRowsBefore;
                if (wordParagraphIndex < 0 || wordParagraphIndex >= paragraphCount)
                    throw new InvalidDataException("分析位置超出 Word 正文段落范围，已停止定位。");
                Word.Range paragraphRange = document.Paragraphs[wordParagraphIndex + 1].Range.Duplicate;
                string paragraphText = paragraphRange.Text ?? string.Empty;
                paragraphText = paragraphText.TrimEnd('\r', '\a');
                string xmlText = string.Concat(paragraphRuns.Select(run => run.Text));
                if (!string.Equals(paragraphText, xmlText, StringComparison.Ordinal)
                    || !string.Equals(xmlParagraph.Text, xmlText, StringComparison.Ordinal))
                    throw new InvalidDataException("第 " + (paragraphIndex + 1) + " 段包含无法安全映射的 Word 内容，未修改文档。");

                int cursor = paragraphRange.Start;
                for (int runIndex = 0; runIndex <= requested.Value.Max(); runIndex++)
                {
                    // Word's internal positions include annotation reference marks,
                    // although Range.Text and OOXML text omit them. Find each next
                    // literal span in order rather than adding string lengths to CPs.
                    Word.Range runRange = FindNextTextRange(document, cursor, paragraphRange.End,
                        paragraphRuns[runIndex].Text);
                    cursor = runRange.End;
                    if (requested.Value.Contains(runIndex))
                        result.Add("word/document.xml#paragraph[" + paragraphIndex + "]/run[" + runIndex + "]", runRange);
                }
            }
            return result;
        }

        private static Word.Range FindNextTextRange(Word.Document document, int cursor, int paragraphEnd, string text)
        {
            if (string.IsNullOrEmpty(text)) return document.Range(cursor, cursor);
            int start = -1;
            int consumed = 0;
            while (consumed < text.Length)
            {
                // Word Find has a short pattern limit. Leave space for escaped carets
                // and never cut a UTF-16 surrogate pair between two literal chunks.
                int length = Math.Min(100, text.Length - consumed);
                if (consumed + length < text.Length && char.IsHighSurrogate(text[consumed + length - 1])) length--;
                string chunk = text.Substring(consumed, length);
                Word.Range target = document.Range(cursor, paragraphEnd);
                Word.Find find = target.Find;
                find.ClearFormatting();
                find.Text = chunk.Replace("^", "^^");
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = true;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                find.MatchFuzzy = false;
                find.IgnoreSpace = false;
                find.IgnorePunct = false;
                if (!find.Execute() || !string.Equals(target.Text, chunk, StringComparison.Ordinal)
                    || !string.IsNullOrEmpty(document.Range(cursor, target.Start).Text))
                    throw new InvalidDataException("无法精确定位原文片段，已停止修改。");
                if (start < 0) start = target.Start;
                cursor = target.End;
                consumed += length;
            }
            Word.Range result = document.Range(start, cursor);
            if (!string.Equals(result.Text, text, StringComparison.Ordinal))
                throw new InvalidDataException("原文片段与分析文本不一致，已停止修改。");
            return result;
        }

        private static Dictionary<int, XmlParagraph> ReadXmlParagraphs(string snapshotPath)
        {
            if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
                throw new FileNotFoundException("用于核对位置的当前文档快照不存在。", snapshotPath);
            using (ZipArchive archive = ZipFile.OpenRead(snapshotPath))
            {
                ZipArchiveEntry entry = archive.GetEntry("word/document.xml");
                if (entry == null) throw new InvalidDataException("当前快照缺少 word/document.xml。");
                using (Stream stream = entry.Open())
                {
                    XDocument document = XDocument.Load(stream, LoadOptions.None);
                    XElement body = document.Root?.Element(W + "body");
                    if (body == null) throw new InvalidDataException("当前快照缺少 Word 正文。");
                    XElement[] bodyRows = body.Descendants(W + "tr")
                        .Where(row => !row.Ancestors(W + "txbxContent").Any())
                        .ToArray();
                    XElement[] orderedNodes = body.DescendantsAndSelf().ToArray();
                    var nodePositions = new Dictionary<XElement, int>();
                    for (int nodeIndex = 0; nodeIndex < orderedNodes.Length; nodeIndex++)
                        nodePositions[orderedNodes[nodeIndex]] = nodeIndex;
                    var rowEndPositions = bodyRows.ToDictionary(row => row,
                        row => nodePositions[row] + row.DescendantsAndSelf().Count() - 1);
                    XElement[] sourceParagraphs = body.Descendants(W + "p")
                        .Where(paragraph => !paragraph.Ancestors(W + "txbxContent").Any())
                        .ToArray();
                    var result = new Dictionary<int, XmlParagraph>();
                    for (int paragraphIndex = 0; paragraphIndex < sourceParagraphs.Length; paragraphIndex++)
                    {
                        XElement sourceParagraph = sourceParagraphs[paragraphIndex];
                        var runs = new List<XmlRun>();
                        foreach (XElement run in sourceParagraph.Elements(W + "r")
                                     .Where(candidate => candidate.Descendants(W + "t").Any()))
                        {
                            string text = string.Concat(run.Descendants(W + "t").Select(element => element.Value));
                            runs.Add(new XmlRun { Text = text });
                        }
                        int paragraphPosition = nodePositions[sourceParagraph];
                        int completedRowsBefore = rowEndPositions.Values.Count(rowEndPosition => rowEndPosition < paragraphPosition);
                        result.Add(paragraphIndex, new XmlParagraph
                        {
                            Runs = runs,
                            Text = string.Concat(sourceParagraph.Descendants(W + "t").Select(element => element.Value)),
                            CompletedTableRowsBefore = completedRowsBefore,
                            TotalTableRows = bodyRows.Length
                        });
                    }
                    return result;
                }
            }
        }

        private static void ValidatePatch(Word.Range range, EnginePatch patch)
        {
            string actual;
            if (string.Equals(patch.RuleId, "demo/east-asian-font", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(patch.DesiredAfter))
                    throw new InvalidDataException("字体修复缺少目标字体。");
                actual = Convert.ToString(range.Font.NameFarEast, CultureInfo.InvariantCulture);
                EnsureBeforeValue(patch, actual);
                return;
            }
            if (string.Equals(patch.RuleId, "demo/western-font", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(patch.DesiredAfter))
                    throw new InvalidDataException("西文字体修复缺少目标字体。");
                actual = Convert.ToString(range.Font.NameAscii, CultureInfo.InvariantCulture);
                EnsureBeforeValue(patch, actual);
                return;
            }
            if (string.Equals(patch.RuleId, "demo/font-size", StringComparison.Ordinal))
            {
                if (!TryParsePoints(patch.DesiredAfter, out _))
                    throw new InvalidDataException("字号修复目标值无效。");
                actual = Convert.ToSingle(range.Font.Size, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture) + "pt";
                EnsureBeforeValue(patch, actual, numeric: true);
                return;
            }

            if (string.Equals(patch.RuleId, "demo/caption-east-asian-font", StringComparison.Ordinal))
            {
                if (!string.Equals(patch.DesiredAfter, "宋体", StringComparison.Ordinal))
                    throw new InvalidDataException("题注字体目标值无效。");
                actual = Convert.ToString(range.Font.NameFarEast, CultureInfo.InvariantCulture);
                EnsureBeforeValue(patch, actual);
                return;
            }
            if (string.Equals(patch.RuleId, "demo/caption-font-size", StringComparison.Ordinal))
            {
                if (!TryParsePoints(patch.DesiredAfter, out _))
                    throw new InvalidDataException("题注字号目标值无效。");
                actual = Convert.ToSingle(range.Font.Size, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture) + "pt";
                EnsureBeforeValue(patch, actual, numeric: true);
                return;
            }
            if (string.Equals(patch.RuleId, "demo/caption-alignment", StringComparison.Ordinal))
            {
                if (!string.Equals(patch.DesiredAfter, "center", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("题注对齐目标值无效。");
                actual = Convert.ToInt32(range.ParagraphFormat.Alignment, CultureInfo.InvariantCulture) ==
                    (int)Word.WdParagraphAlignment.wdAlignParagraphCenter ? "center" : "not-center";
                EnsureBeforeValue(patch, actual);
                return;
            }
            if (string.Equals(patch.RuleId, "demo/caption-line-spacing", StringComparison.Ordinal))
            {
                if (!string.Equals(patch.DesiredAfter, "single", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("题注行距目标值无效。");
                actual = Convert.ToInt32(range.ParagraphFormat.LineSpacingRule, CultureInfo.InvariantCulture) ==
                    (int)Word.WdLineSpacing.wdLineSpaceSingle ? "single" : "not-single";
                EnsureBeforeValue(patch, actual);
                return;
            }
            if (string.Equals(patch.RuleId, "demo/caption-space-before", StringComparison.Ordinal) ||
                string.Equals(patch.RuleId, "demo/caption-space-after", StringComparison.Ordinal))
            {
                if (!TryParsePoints(patch.DesiredAfter, out _))
                    throw new InvalidDataException("题注段前/段后距目标值无效。");
                float points = string.Equals(patch.RuleId, "demo/caption-space-before", StringComparison.Ordinal)
                    ? range.ParagraphFormat.SpaceBefore : range.ParagraphFormat.SpaceAfter;
                actual = points.ToString(CultureInfo.InvariantCulture) + "pt";
                EnsureBeforeValue(patch, actual, numeric: true);
                return;
            }
            throw new NotSupportedException("Word 当前不支持此规则的直接修复：" + patch.RuleId);
        }

        private static void EnsureBeforeValue(EnginePatch patch, string actual, bool numeric = false)
        {
            bool matches = numeric && TryParsePoints(actual, out double actualPoints) &&
                TryParsePoints(patch.ExpectedBefore, out double expectedPoints)
                    ? Math.Abs(actualPoints - expectedPoints) <= 0.01
                    : string.Equals(actual, patch.ExpectedBefore, StringComparison.Ordinal);
            if (!matches)
                throw new InvalidDataException("修复前置条件已变化，已停止本次批量操作：期望「" +
                    patch.ExpectedBefore + "」，当前为「" + actual + "」。");
        }

        private static void ApplyPatch(Word.Range range, EnginePatch patch)
        {
            if (string.Equals(patch.RuleId, "demo/east-asian-font", StringComparison.Ordinal))
                range.Font.NameFarEast = patch.DesiredAfter;
            else if (string.Equals(patch.RuleId, "demo/western-font", StringComparison.Ordinal))
            {
                range.Font.NameAscii = patch.DesiredAfter;
                range.Font.NameOther = patch.DesiredAfter;
            }
            else if (string.Equals(patch.RuleId, "demo/font-size", StringComparison.Ordinal) && TryParsePoints(patch.DesiredAfter, out double points))
                range.Font.Size = (float)points;
            else if (string.Equals(patch.RuleId, "demo/caption-east-asian-font", StringComparison.Ordinal))
                range.Font.NameFarEast = patch.DesiredAfter;
            else if (string.Equals(patch.RuleId, "demo/caption-font-size", StringComparison.Ordinal) && TryParsePoints(patch.DesiredAfter, out double captionSize))
                range.Font.Size = (float)captionSize;
            else if (string.Equals(patch.RuleId, "demo/caption-alignment", StringComparison.Ordinal))
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            else if (string.Equals(patch.RuleId, "demo/caption-line-spacing", StringComparison.Ordinal))
                range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            else if (string.Equals(patch.RuleId, "demo/caption-space-before", StringComparison.Ordinal) && TryParsePoints(patch.DesiredAfter, out double before))
                range.ParagraphFormat.SpaceBefore = (float)before;
            else if (string.Equals(patch.RuleId, "demo/caption-space-after", StringComparison.Ordinal) && TryParsePoints(patch.DesiredAfter, out double after))
                range.ParagraphFormat.SpaceAfter = (float)after;
        }

        private static bool TryParsePoints(string value, out double points)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                normalized = normalized.Substring(0, normalized.Length - 2).Trim();
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out points) && points > 0;
        }

        private static IEnumerable<string> GetDiagnosticLocations(EngineAnalysisResult analysis)
        {
            foreach (EngineDiagnostic diagnostic in analysis?.Diagnostics ?? new List<EngineDiagnostic>())
                foreach (string location in GetOccurrences(diagnostic))
                    yield return location;
        }

        private static IEnumerable<string> GetOccurrences(EngineDiagnostic diagnostic)
        {
            if (diagnostic.Occurrences != null && diagnostic.Occurrences.Count > 0)
                return diagnostic.Occurrences.Select(occurrence => occurrence.Location).Distinct(StringComparer.Ordinal);
            return new[] { diagnostic.Location };
        }

        private static string GetHighlight(Word.Range range)
        {
            return GetHighlightName((Word.WdColorIndex)range.HighlightColorIndex);
        }

        private static string GetHighlightName(Word.WdColorIndex value)
        {
            switch (value)
            {
                case Word.WdColorIndex.wdNoHighlight: return "absent";
                case Word.WdColorIndex.wdBlack: return "black";
                case Word.WdColorIndex.wdBlue: return "blue";
                case Word.WdColorIndex.wdTurquoise: return "cyan";
                case Word.WdColorIndex.wdBrightGreen: return "green";
                case Word.WdColorIndex.wdPink: return "magenta";
                case Word.WdColorIndex.wdRed: return "red";
                case Word.WdColorIndex.wdYellow: return "yellow";
                case Word.WdColorIndex.wdWhite: return "white";
                case Word.WdColorIndex.wdDarkBlue: return "darkBlue";
                case Word.WdColorIndex.wdTeal: return "darkCyan";
                case Word.WdColorIndex.wdGreen: return "darkGreen";
                case Word.WdColorIndex.wdViolet: return "darkMagenta";
                case Word.WdColorIndex.wdDarkRed: return "darkRed";
                case Word.WdColorIndex.wdDarkYellow: return "darkYellow";
                case Word.WdColorIndex.wdGray50: return "darkGray";
                case Word.WdColorIndex.wdGray25: return "lightGray";
                default: throw new InvalidDataException("原文片段使用了混合或不支持的高亮，未刷新批注。");
            }
        }

        private static void SetHighlight(Word.Range range, string value)
        {
            switch (value)
            {
                case "absent": range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight; break;
                case "black": range.HighlightColorIndex = Word.WdColorIndex.wdBlack; break;
                case "blue": range.HighlightColorIndex = Word.WdColorIndex.wdBlue; break;
                case "cyan": range.HighlightColorIndex = Word.WdColorIndex.wdTurquoise; break;
                case "green": range.HighlightColorIndex = Word.WdColorIndex.wdBrightGreen; break;
                case "magenta": range.HighlightColorIndex = Word.WdColorIndex.wdPink; break;
                case "red": range.HighlightColorIndex = Word.WdColorIndex.wdRed; break;
                case "yellow": range.HighlightColorIndex = Word.WdColorIndex.wdYellow; break;
                case "white": range.HighlightColorIndex = Word.WdColorIndex.wdWhite; break;
                case "darkBlue": range.HighlightColorIndex = Word.WdColorIndex.wdDarkBlue; break;
                case "darkCyan": range.HighlightColorIndex = Word.WdColorIndex.wdTeal; break;
                case "darkGreen": range.HighlightColorIndex = Word.WdColorIndex.wdGreen; break;
                case "darkMagenta": range.HighlightColorIndex = Word.WdColorIndex.wdViolet; break;
                case "darkRed": range.HighlightColorIndex = Word.WdColorIndex.wdDarkRed; break;
                case "darkYellow": range.HighlightColorIndex = Word.WdColorIndex.wdDarkYellow; break;
                case "darkGray": range.HighlightColorIndex = Word.WdColorIndex.wdGray50; break;
                case "lightGray": range.HighlightColorIndex = Word.WdColorIndex.wdGray25; break;
                default: throw new InvalidDataException("无法恢复原文高亮值：" + value);
            }
        }
    }
}
