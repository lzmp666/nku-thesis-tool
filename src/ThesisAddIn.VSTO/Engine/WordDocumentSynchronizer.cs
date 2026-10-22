using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ThesisAddIn.Protocol;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>
    /// Applies engine patches and refreshes NKU-owned comments in the open Word document.
    /// No document is saved or reopened here, so Word keeps the edits in its normal undo stack
    /// and the author decides when to save the working copy.
    /// </summary>
    public static class WordDocumentSynchronizer
    {
        private const string ToolAuthor = "NKU Thesis Tool";
        private const string ToolMarker = "[NKU-TOOL-ANNOTATION:v1]";
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

        private sealed class XmlRun
        {
            public string Text { get; set; }
            public int Offset { get; set; }
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
            return ExecuteUndoable(document, "刷新论文格式批注", () =>
            {
                RefreshAnnotationsCore(document, snapshotPath, analysis);
                document.Saved = false;
            }, out error);
        }

        /// <summary>Applies accepted patches and refreshes marks without replacing the open document.</summary>
        public static bool ApplyRepairsAndRefresh(Word.Document document, string snapshotPath,
            IReadOnlyList<EnginePatch> patches, EngineAnalysisResult analysis, out string error)
        {
            if (patches == null) throw new ArgumentNullException(nameof(patches));
            return ExecuteUndoable(document, "修复论文格式", () =>
            {
                Dictionary<string, Word.Range> ranges = BuildRanges(document, snapshotPath,
                    patches.Select(patch => patch.Location)
                        .Concat(GetDiagnosticLocations(analysis)));

                // Reject unsupported or stale work before changing the document.
                foreach (EnginePatch patch in patches)
                {
                    if (patch == null || !ranges.ContainsKey(patch.Location))
                        throw new InvalidDataException("修复计划中的位置无法映射到当前 Word 文本。");
                    ValidatePatch(patch);
                }

                foreach (EnginePatch patch in patches)
                    ApplyPatch(ranges[patch.Location], patch);

                RefreshAnnotationsCore(document, snapshotPath, analysis);
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

        private static bool ExecuteUndoable(Word.Document document, string actionName, Action action, out string error)
        {
            error = null;
            if (document == null)
            {
                error = "Word 文档已关闭，请重新打开后重试。";
                return false;
            }

            dynamic undoRecord = null;
            bool customRecordStarted = false;
            string before = null;
            try
            {
                before = document.WordOpenXML;
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

                action();
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
                        string after = document.WordOpenXML;
                        if (!string.Equals(before, after, StringComparison.Ordinal))
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

        private static void RefreshAnnotationsCore(Word.Document document, string snapshotPath,
            EngineAnalysisResult analysis)
        {
            if (analysis == null) throw new ArgumentNullException(nameof(analysis));

            List<AnnotationRecord> existing = ReadOwnedAnnotations(document);
            var locations = new HashSet<string>(StringComparer.Ordinal);
            foreach (AnnotationRecord record in existing)
                foreach (string location in record.OriginalHighlights.Keys)
                    locations.Add(location);
            foreach (string location in GetDiagnosticLocations(analysis))
                locations.Add(location);

            Dictionary<string, Word.Range> ranges = BuildRanges(document, snapshotPath, locations);
            // Keep the work-copy identity after the last diagnostic/comment is removed.
            // Word persists document variables when the author saves the working copy.
            MarkWorkingCopy(document);

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
                SetHighlight(ranges[highlight.Key], highlight.Value);

            // Delete in reverse order so the 1-based collection indexes remain valid.
            foreach (int index in existing.Select(record => record.Index).OrderByDescending(index => index))
                document.Comments[index].Delete();

            // Comment removal can shift story ranges. Rebuild all targets against the same
            // text snapshot before writing new highlights and anchors.
            string[] newLocations = GetDiagnosticLocations(analysis).ToArray();
            ranges = BuildRanges(document, snapshotPath, newLocations);
            var freshOriginalHighlights = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string location in newLocations.Distinct(StringComparer.Ordinal))
                freshOriginalHighlights[location] = GetHighlight(ranges[location]);

            foreach (EngineDiagnostic diagnostic in analysis.Diagnostics ?? new List<EngineDiagnostic>())
            {
                string[] occurrences = GetOccurrences(diagnostic).ToArray();
                foreach (string location in occurrences)
                    ranges[location].HighlightColorIndex = Word.WdColorIndex.wdYellow;

                string visibleText = "要求参考来源：2026版第4.7节表4.6；规则：" + diagnostic.RuleId
                    + "；期望值：" + diagnostic.Expected + "；实际值：" + diagnostic.Actual + "。";
                string hiddenText = BuildHiddenMetadata(diagnostic.Location, freshOriginalHighlights[diagnostic.Location],
                    occurrences, freshOriginalHighlights);
                AddOwnedComment(document, ranges[diagnostic.Location], visibleText, hiddenText);
            }
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

        private static Dictionary<string, Word.Range> BuildRanges(Word.Document document, string snapshotPath,
            IEnumerable<string> requestedLocations)
        {
            if (document.Tables.Count > 0)
                throw new NotSupportedException("当前定位器不支持含表格的文档，未修改内容。");

            Dictionary<int, List<XmlRun>> runsByParagraph = ReadXmlRuns(snapshotPath);
            int paragraphCount = document.Paragraphs.Count;
            int xmlParagraphCount = runsByParagraph.Count == 0 ? 0 : runsByParagraph.Keys.Max() + 1;
            if (paragraphCount != xmlParagraphCount)
                throw new InvalidDataException("Word 段落与分析快照无法一一对应，为避免改错位置已停止。");

            var result = new Dictionary<string, Word.Range>(StringComparer.Ordinal);
            foreach (string location in requestedLocations.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal))
            {
                Match match = LocationPattern.Match(location);
                if (!match.Success)
                    throw new InvalidDataException("暂不支持该位置格式：" + location);
                int paragraphIndex = int.Parse(match.Groups["paragraph"].Value, CultureInfo.InvariantCulture);
                int runIndex = int.Parse(match.Groups["run"].Value, CultureInfo.InvariantCulture);
                if (!runsByParagraph.TryGetValue(paragraphIndex, out List<XmlRun> paragraphRuns) ||
                    runIndex < 0 || runIndex >= paragraphRuns.Count)
                    throw new InvalidDataException("分析位置已不存在：" + location);

                Word.Range paragraphRange = document.Paragraphs[paragraphIndex + 1].Range.Duplicate;
                string paragraphText = paragraphRange.Text ?? string.Empty;
                paragraphText = paragraphText.TrimEnd('\r', '\a');
                string xmlText = string.Concat(paragraphRuns.Select(run => run.Text));
                if (!string.Equals(paragraphText, xmlText, StringComparison.Ordinal))
                    throw new InvalidDataException("第 " + (paragraphIndex + 1) + " 段包含无法安全映射的 Word 内容，未修改文档。");

                int offset = paragraphRuns[runIndex].Offset;
                string expectedText = paragraphRuns[runIndex].Text;
                int start = paragraphRange.Start + offset;
                Word.Range runRange = document.Range(start, start + expectedText.Length);
                if (!string.Equals(runRange.Text, expectedText, StringComparison.Ordinal))
                    throw new InvalidDataException("无法精确定位原文片段：" + location);
                result.Add(location, runRange);
            }
            return result;
        }

        private static Dictionary<int, List<XmlRun>> ReadXmlRuns(string snapshotPath)
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
                    var result = new Dictionary<int, List<XmlRun>>();
                    XElement[] paragraphs = body.Elements(W + "p").ToArray();
                    for (int paragraphIndex = 0; paragraphIndex < paragraphs.Length; paragraphIndex++)
                    {
                        int offset = 0;
                        var runs = new List<XmlRun>();
                        foreach (XElement run in paragraphs[paragraphIndex].Elements(W + "r")
                                     .Where(candidate => candidate.Descendants(W + "t").Any()))
                        {
                            string text = string.Concat(run.Descendants(W + "t").Select(element => element.Value));
                            runs.Add(new XmlRun
                            {
                                Text = text,
                                Offset = offset
                            });
                            offset += text.Length;
                        }
                        result.Add(paragraphIndex, runs);
                    }
                    return result;
                }
            }
        }

        private static void ValidatePatch(EnginePatch patch)
        {
            if (string.Equals(patch.RuleId, "demo/east-asian-font", StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(patch.DesiredAfter))
                    throw new InvalidDataException("字体修复缺少目标字体。");
                return;
            }
            if (string.Equals(patch.RuleId, "demo/font-size", StringComparison.Ordinal))
            {
                if (!TryParsePoints(patch.DesiredAfter, out _))
                    throw new InvalidDataException("字号修复目标值无效。");
                return;
            }
            throw new NotSupportedException("Word 当前不支持此规则的直接修复：" + patch.RuleId);
        }

        private static void ApplyPatch(Word.Range range, EnginePatch patch)
        {
            if (string.Equals(patch.RuleId, "demo/east-asian-font", StringComparison.Ordinal))
                range.Font.NameFarEast = patch.DesiredAfter;
            else if (string.Equals(patch.RuleId, "demo/font-size", StringComparison.Ordinal) && TryParsePoints(patch.DesiredAfter, out double points))
                range.Font.Size = (float)points;
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

        private static string BuildHiddenMetadata(string primaryLocation, string primaryHighlight,
            IEnumerable<string> occurrences, IDictionary<string, string> originalHighlights)
        {
            var text = new StringBuilder(ToolMarker)
                .Append("[NKU-LOCATION:").Append(primaryLocation).Append(']')
                .Append("[NKU-HIGHLIGHT:").Append(primaryHighlight).Append(']');
            foreach (string location in occurrences.Distinct(StringComparer.Ordinal))
                text.Append("[NKU-OCCURRENCE:").Append(location).Append('|')
                    .Append(originalHighlights[location]).Append(']');
            return text.ToString();
        }

        private static void AddOwnedComment(Word.Document document, Word.Range anchor, string visibleText, string hiddenText)
        {
            object text = visibleText;
            Word.Comment comment = document.Comments.Add(anchor, ref text);
            // Signed-in Word can ignore Application.UserName when creating a comment.
            // Set the comment itself so both Word and the engine recognize its ownership.
            comment.Author = ToolAuthor;
            comment.Initial = "NKU";
            Word.Range metadata = comment.Range.Duplicate;
            metadata.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            int start = metadata.End;
            metadata.InsertAfter(hiddenText);
            metadata.SetRange(start, start + hiddenText.Length);
            metadata.Font.Hidden = -1;
        }

        private static string GetHighlight(Word.Range range)
        {
            var value = (Word.WdColorIndex)range.HighlightColorIndex;
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
