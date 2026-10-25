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
    public static partial class WordDocumentSynchronizer
    {
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
            public Word.Comment Comment { get; set; }
            public int Index { get; set; }
            public string Text { get; set; }
            public Word.Range Scope { get; set; }
            public Dictionary<string, string> OriginalHighlights { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
        }

        private sealed class HighlightBookmarkRecord
        {
            public string Name { get; set; }
            public int OriginalHighlight { get; set; }
            public Word.Range Range { get; set; }
        }

        private sealed class AnnotationRefreshState
        {
            public List<AnnotationRecord> Comments;
            public List<HighlightBookmarkRecord> Marks;
            public Dictionary<string, Word.Range> Ranges;
            public Dictionary<string, string> Texts;
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
            EngineAnalysisResult analysis, out string error, Action<string> trace = null)
        {
            Dictionary<string, Word.Range> ranges;
            return RefreshAnnotations(document, snapshotPath, analysis, out error, out ranges, trace);
        }

        public static bool RefreshAnnotations(Word.Document document, string snapshotPath,
            EngineAnalysisResult analysis, out string error, out Dictionary<string, Word.Range> ranges,
            Action<string> trace = null, Dictionary<string, Word.Range> cached = null)
        {
            AnnotationRefreshState prepared = null;
            Dictionary<string, Word.Range> updated = null;
            bool success = ExecuteUndoable(document, "刷新论文格式批注",
                () => prepared = PrepareAnnotationRefresh(document, snapshotPath, analysis, cached), markMutation =>
            {
                updated = RefreshAnnotationsCore(document, snapshotPath, analysis, markMutation, trace, prepared);
                document.Saved = false;
            }, out error);
            ranges = success ? updated : null;
            return success;
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
                    ApplyPatch(ranges[patch.Location], patch, markMutation);
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
            bool? screenUpdating = null;
            try
            {
                // Suspend repaint only during synchronous COM work, never across an await.
                // Restoring the caller's value also covers validation and rollback failures.
                screenUpdating = document.Application.ScreenUpdating;
                document.Application.ScreenUpdating = false;
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
                if (screenUpdating.HasValue)
                {
                    try { document.Application.ScreenUpdating = screenUpdating.Value; }
                    catch (System.Runtime.InteropServices.COMException) { }
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

        private static AnnotationRefreshState PrepareAnnotationRefresh(Word.Document document, string snapshotPath,
            EngineAnalysisResult analysis, Dictionary<string, Word.Range> cached = null)
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

            Dictionary<string, Word.Range> ranges = BuildRanges(document, snapshotPath, locations, cached);
            return new AnnotationRefreshState { Comments = existing, Marks = existingMarks, Ranges = ranges,
                Texts = ranges.ToDictionary(pair => pair.Key, pair => pair.Value.Text, StringComparer.Ordinal) };
        }

        private static Dictionary<string, Word.Range> RefreshAnnotationsCore(Word.Document document, string snapshotPath,
            EngineAnalysisResult analysis, Action markMutation, Action<string> trace = null,
            AnnotationRefreshState prepared = null)
        {
            prepared = prepared ?? PrepareAnnotationRefresh(document, snapshotPath, analysis);
            List<AnnotationRecord> existing = prepared.Comments;
            List<HighlightBookmarkRecord> existingMarks = prepared.Marks;
            Dictionary<string, Word.Range> ranges = prepared.Ranges;
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

            string[] newLocations = GetDiagnosticLocations(analysis).Distinct(StringComparer.Ordinal).ToArray();
            var wantedLocations = new HashSet<string>(newLocations, StringComparer.Ordinal);
            var freshOriginalHighlights = new Dictionary<string, int>(StringComparer.Ordinal);
            // Recover the pre-tool color without clearing/reapplying every yellow range.
            foreach (string location in newLocations)
            {
                string original;
                if (originalByLocation.TryGetValue(location, out original))
                    freshOriginalHighlights[location] = (int)HighlightValue(original);
                else
                {
                    _ = GetHighlight(ranges[location]);
                    freshOriginalHighlights[location] = (int)ranges[location].HighlightColorIndex;
                }
            }
            var retainedComments = new HashSet<int>();
            var retainedDiagnostics = new HashSet<EngineDiagnostic>();
            foreach (EngineDiagnostic diagnostic in analysis.Diagnostics ?? new List<EngineDiagnostic>())
            {
                string[] occurrences = GetOccurrences(diagnostic).ToArray();
                if (!CreateNativeComments || occurrences.Length == 0) continue;
                string desired = NormalizeCommentText(BuildVisibleCommentText(diagnostic) +
                    BuildHiddenCommentMetadata(diagnostic, occurrences[0], occurrences, freshOriginalHighlights));
                Word.Range anchor = ranges[occurrences[0]];
                AnnotationRecord match = existing.FirstOrDefault(record => !retainedComments.Contains(record.Index)
                    && NormalizeCommentText(record.Text) == desired
                    && record.Scope.Start == anchor.Start && record.Scope.End == anchor.End);
                if (match == null) continue;
                retainedComments.Add(match.Index);
                retainedDiagnostics.Add(diagnostic);
            }
            foreach (KeyValuePair<string, string> highlight in originalByLocation)
            {
                if (wantedLocations.Contains(highlight.Key)) continue;
                if (string.Equals(GetHighlight(ranges[highlight.Key]), highlight.Value, StringComparison.Ordinal)) continue;
                SetHighlight(ranges[highlight.Key], highlight.Value);
                markMutation();
            }

            var retainedMarks = new HashSet<string>(StringComparer.Ordinal);
            foreach (string location in newLocations)
            {
                string name = HighlightBookmarkName(location, freshOriginalHighlights[location]);
                HighlightBookmarkRecord mark = existingMarks.FirstOrDefault(candidate => candidate.Name == name);
                if (mark?.Range != null && mark.Range.Start == ranges[location].Start && mark.Range.End == ranges[location].End)
                    retainedMarks.Add(name);
            }
            foreach (HighlightBookmarkRecord mark in existingMarks)
            {
                if (retainedMarks.Contains(mark.Name)) continue;
                if (mark.Range == null) continue;
                if ((int)mark.Range.HighlightColorIndex != mark.OriginalHighlight)
                {
                    mark.Range.HighlightColorIndex = (Word.WdColorIndex)mark.OriginalHighlight;
                    markMutation();
                }
            }

            // Delete in reverse order so the 1-based collection indexes remain valid.
            foreach (int index in existing.Where(record => !retainedComments.Contains(record.Index))
                .Select(record => record.Index).OrderByDescending(index => index))
            {
                document.Comments[index].Delete();
                markMutation();
            }

            foreach (HighlightBookmarkRecord mark in existingMarks)
            {
                if (retainedMarks.Contains(mark.Name)) continue;
                if (!document.Bookmarks.Exists(mark.Name)) continue;
                document.Bookmarks[mark.Name].Delete();
                markMutation();
            }

            // Live Word.Range objects follow deleted/inserted comment reference marks.
            // Verify their text again rather than repeating XML parsing and COM Find.
            foreach (string location in newLocations)
                if (!string.Equals(ranges[location].Text, prepared.Texts[location], StringComparison.Ordinal))
                    throw new InvalidDataException("批注更新期间原文发生变化，已停止操作。");

            var markedLocations = new HashSet<string>(StringComparer.Ordinal);
            foreach (EngineDiagnostic diagnostic in analysis.Diagnostics ?? new List<EngineDiagnostic>())
            {
                string[] occurrences = GetOccurrences(diagnostic).ToArray();
                foreach (string location in occurrences)
                {
                    if (markedLocations.Add(location) && !retainedMarks.Contains(
                        HighlightBookmarkName(location, freshOriginalHighlights[location])))
                        AddHighlightBookmark(document, location, ranges[location],
                            freshOriginalHighlights[location], markMutation);
                    if (ranges[location].HighlightColorIndex != Word.WdColorIndex.wdYellow)
                    {
                        ranges[location].HighlightColorIndex = Word.WdColorIndex.wdYellow;
                        markMutation();
                    }
                }

                // Keep one native Word comment for each grouped issue. The comment is
                // anchored to the first occurrence; every occurrence remains highlighted
                // and the hidden location metadata keeps the issue list and hover in sync.
                if (occurrences.Length > 0 && CreateNativeComments && !retainedDiagnostics.Contains(diagnostic))
                {
                    string anchorLocation = occurrences[0];
                    string visibleText = BuildVisibleCommentText(diagnostic);
                    string hiddenMetadata = BuildHiddenCommentMetadata(diagnostic, anchorLocation,
                        occurrences, freshOriginalHighlights);
                    Word.Range anchor = ranges[anchorLocation].Duplicate;
                    object commentText = visibleText + hiddenMetadata;
                    Word.Comment comment = document.Comments.Add(anchor, ref commentText);
                    markMutation();
                    try
                    {
                        comment.Author = "NKU Thesis Tool";
                        comment.Initial = "NT";
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Some Word comment providers retain the signed-in author.
                        // Ownership is identified by the private marker, not the author.
                    }

                    // The private markers are needed for refresh, hover and safe cleanup,
                    // but must never appear in the user's visible comment.
                    Word.Range hiddenRange = comment.Range.Duplicate;
                    string insertedText = hiddenRange.Text ?? string.Empty;
                    int markerOffset = insertedText.IndexOf(ToolMarker, StringComparison.Ordinal);
                    if (markerOffset < 0)
                        throw new InvalidDataException("Word 未保留问题批注的位置标记。");
                    int metadataStart = hiddenRange.Start + markerOffset;
                    StyleComment(comment, insertedText, metadataStart);
                    hiddenRange.SetRange(metadataStart, metadataStart + hiddenMetadata.Length);
                    hiddenRange.Font.Hidden = -1;
                }
            }
            // This legacy marker identifies an analyzed document, not a separate copy.
            // Write it last; document variables do not create Word undo entries.
            MarkWorkingCopy(document);
            if (CreateNativeComments && document.Comments.Count > 0) ShowCommentBalloons(document);
            else RestoreCommentWidth(document);
            trace?.Invoke("批注增量更新 retained=" + retainedComments.Count + " removed=" +
                (existing.Count - retainedComments.Count) + " added=" +
                ((analysis.Diagnostics?.Count ?? 0) - retainedDiagnostics.Count) +
                " highlightsRetained=" + retainedMarks.Count);
            return newLocations.ToDictionary(location => location, location => ranges[location], StringComparer.Ordinal);
        }

        private static void StyleComment(Word.Comment comment, string text, int metadataStart)
        {
            Word.Range body = comment.Range.Duplicate;
            body.SetRange(body.Start, metadataStart);
            body.Font.Name = "Microsoft YaHei";
            body.Font.Size = 9.5F;
            body.Font.Bold = 0;
            body.Font.Color = (Word.WdColor)0x2A272C;
            body.ParagraphFormat.SpaceBefore = 0;
            body.ParagraphFormat.SpaceAfter = 3;
            body.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            foreach (string label in new[] { "要求：", "实际：" })
            {
                int offset = text.IndexOf(label, StringComparison.Ordinal);
                if (offset < 0) continue;
                Word.Range field = body.Duplicate;
                field.SetRange(body.Start + offset, body.Start + offset + label.Length);
                field.Font.Bold = -1;
                field.Font.Color = (Word.WdColor)0x3D1F68; // Nankai burgundy, Word's BGR order.
            }
        }

        private static void ShowCommentBalloons(Word.Document document)
        {
            try
            {
                Word.View view = document.ActiveWindow.View;
                view.ShowComments = true;
                view.ShowRevisionsAndComments = true;
                view.RevisionsMode = Word.WdRevisionsMode.wdBalloonRevisions;
                view.RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupAll;
                SetCompactCommentWidth(view);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // The annotations remain intact when this Word version cannot change its view.
            }
        }

        private static float? _previousBalloonWidth;
        private static Word.WdRevisionsBalloonWidthType _previousBalloonWidthType;
        private static void SetCompactCommentWidth(Word.View view)
        {
            // Word documents these as application-global settings. Capture once and
            // restore on switching to another document, closing this one, or shutdown.
            if (!_previousBalloonWidth.HasValue)
            {
                _previousBalloonWidth = view.RevisionsBalloonWidth;
                _previousBalloonWidthType = view.RevisionsBalloonWidthType;
            }
            view.RevisionsBalloonWidthType = Word.WdRevisionsBalloonWidthType.wdBalloonWidthPoints;
            view.RevisionsBalloonWidth = 210F;
        }
        public static void UpdateCommentWidth(Word.Document document)
        {
            try
            {
                if (IsWorkingCopy(document) && ReadOwnedAnnotations(document).Count > 0)
                    SetCompactCommentWidth(document.ActiveWindow.View);
                else RestoreCommentWidth(document);
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }
        public static void RestoreCommentWidth(Word.Document document)
        {
            if (!_previousBalloonWidth.HasValue || document == null) return;
            try
            {
                Word.View view = document.ActiveWindow.View;
                view.RevisionsBalloonWidthType = _previousBalloonWidthType;
                view.RevisionsBalloonWidth = _previousBalloonWidth.Value;
                _previousBalloonWidth = null;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        private static List<AnnotationRecord> ReadOwnedAnnotations(Word.Document document, Func<string,bool> keep = null)
        {
            var records = new List<AnnotationRecord>();
            Word.Comments comments = document.Comments;
            int count = comments.Count;
            for (int index = 1; index <= count; index++)
            {
                Word.Comment comment = comments[index];
                Word.Range range = comment.Range.Duplicate;
                range.TextRetrievalMode.IncludeHiddenText = true;
                string text = range.Text ?? string.Empty;
                if (text.IndexOf(ToolMarker, StringComparison.Ordinal) < 0)
                    continue;
                if (keep != null && !keep(text)) continue;

                var record = new AnnotationRecord { Comment = comment, Index = index, Text = text, Scope = comment.Scope.Duplicate };
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

        private static string BuildVisibleCommentText(EngineDiagnostic diagnostic)
        {
            var fields = new List<string>();
            fields.Add("要求：" + DisplayCommentValue(diagnostic.Expected));
            fields.Add("实际：" + DisplayCommentValue(diagnostic.Actual));
            return string.Join(Environment.NewLine, fields) + "。";
        }

        private static string NormalizeCommentText(string text) =>
            (text ?? string.Empty).Replace("\r\n", "\r").Replace("\n", "\r").TrimEnd('\r');

        private static Word.WdColorIndex HighlightValue(string name)
        {
            foreach (Word.WdColorIndex value in Enum.GetValues(typeof(Word.WdColorIndex)))
            {
                try { if (GetHighlightName(value) == name) return value; }
                catch (InvalidDataException) { }
            }
            throw new InvalidDataException("无法恢复原文高亮值：" + name);
        }

        private static string DisplayCommentValue(string value)
        {
            switch(value) {
                case "center": return "居中"; case "left": return "左对齐";
                case "single": return "单倍行距"; case "multiple:360": return "1.5 倍行距";
                case "(unknown)": return "未知，需核对";
                default: return value ?? "未知";
            }
        }

        private static string GetSourceReference(EngineDiagnostic diagnostic)
        {
            string message = diagnostic?.Message ?? string.Empty;
            if ((diagnostic?.RuleId ?? string.Empty).StartsWith("NKU-REQ-", StringComparison.Ordinal))
            {
                int separator = message.IndexOf('：');
                if (separator > 0) return message.Substring(0, separator);
            }

            const string sourcePrefix = "依据：";
            if (message.StartsWith(sourcePrefix, StringComparison.Ordinal))
            {
                int end = message.IndexOf('。', sourcePrefix.Length);
                if (end > sourcePrefix.Length) return message.Substring(sourcePrefix.Length, end - sourcePrefix.Length);
            }

            return (diagnostic?.RuleId ?? string.Empty).StartsWith("demo/", StringComparison.Ordinal)
                ? "演示规则（仅用于验证插件流程）" : string.Empty;
        }

        private static string BuildHiddenCommentMetadata(EngineDiagnostic diagnostic, string anchorLocation,
            IEnumerable<string> locations, IDictionary<string, int> originalHighlights)
        {
            var metadata = new StringBuilder();
            metadata.Append(ToolMarker);
            // Retain the stable machine fields as hidden metadata for source linking.
            metadata.Append("\r规则：").Append(diagnostic.RuleId).Append("\r期望值：").Append(diagnostic.Expected)
                .Append("\r实际值：").Append(diagnostic.Actual).Append("\r");
            metadata.Append("[NKU-LOCATION:").Append(anchorLocation).Append(']');
            if (originalHighlights.TryGetValue(anchorLocation, out int original))
                metadata.Append("[NKU-HIGHLIGHT:").Append(GetHighlightName((Word.WdColorIndex)original)).Append(']');

            foreach (string location in locations.Distinct(StringComparer.Ordinal))
            {
                if (!originalHighlights.TryGetValue(location, out int highlight))
                    throw new InvalidDataException("问题批注缺少原始高亮元数据：" + location);
                metadata.Append("[NKU-OCCURRENCE:").Append(location).Append('|')
                    .Append(GetHighlightName((Word.WdColorIndex)highlight)).Append(']');
            }
            return metadata.ToString();
        }

        private static List<HighlightBookmarkRecord> ReadOwnedHighlightBookmarks(Word.Document document)
        {
            var records = new List<HighlightBookmarkRecord>();
            Word.Bookmarks bookmarks = document.Bookmarks;
            bool previousShowHidden = bookmarks.ShowHidden;
            // Our underscore-prefixed recovery bookmarks are hidden. Word omits
            // them from Count/index enumeration unless ShowHidden is enabled.
            // Restore the caller's setting after taking the owned-range snapshot.
            try
            {
                bookmarks.ShowHidden = true;
                for (int index = bookmarks.Count; index >= 1; index--)
                {
                    Word.Bookmark bookmark = bookmarks[index];
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
            }
            finally { bookmarks.ShowHidden = previousShowHidden; }
            return records;
        }

        private static void AddHighlightBookmark(Word.Document document, string location, Word.Range range,
            int originalHighlight, Action markMutation)
        {
            string name = HighlightBookmarkName(location, originalHighlight);
            if (document.Bookmarks.Exists(name))
                throw new InvalidDataException("插件的高亮恢复标记已存在；请重新分析后再操作。");
            object bookmarkRange = range.Duplicate;
            document.Bookmarks.Add(name, ref bookmarkRange);
            markMutation();
        }

        private static string HighlightBookmarkName(string location, int originalHighlight)
        {
            string locationHash;
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(location));
                locationHash = BitConverter.ToString(bytes).Replace("-", string.Empty).Substring(0, 16);
            }
            return ToolBookmarkPrefix + locationHash + "_" +
                originalHighlight.ToString("X8", CultureInfo.InvariantCulture);
        }

        private static Dictionary<string, Word.Range> BuildRanges(Word.Document document, string snapshotPath,
            IEnumerable<string> requestedLocations, Dictionary<string, Word.Range> cached = null)
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
                // A cache hit is accepted only after matching the live paragraph,
                // exact run text AND prefix. Live ranges follow Word's hidden comment
                // reference CPs; offsets calculated from plain string lengths do not.
                bool allCached = cached != null && requested.Value.All(runIndex => {
                    string location = "word/document.xml#paragraph[" + paragraphIndex + "]/run[" + runIndex + "]";
                    Word.Range candidate;
                    if (!cached.TryGetValue(location, out candidate) || candidate == null) return false;
                    try {
                        return candidate.Start >= paragraphRange.Start && candidate.End <= paragraphRange.End &&
                            candidate.End - candidate.Start == paragraphRuns[runIndex].Text.Length &&
                            candidate.Text == paragraphRuns[runIndex].Text &&
                            document.Range(paragraphRange.Start, candidate.Start).Text ==
                                string.Concat(paragraphRuns.Take(runIndex).Select(run => run.Text));
                    } catch (System.Runtime.InteropServices.COMException) { return false; }
                });
                if (allCached) {
                    foreach (int runIndex in requested.Value) {
                        string location = "word/document.xml#paragraph[" + paragraphIndex + "]/run[" + runIndex + "]";
                        result.Add(location, cached[location]);
                    }
                    continue;
                }
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
            // Most runs have no annotation reference inside them. Validate a direct span
            // first; retain literal Find for Word CP offsets containing hidden references.
            if (text.Length <= paragraphEnd - cursor)
            {
                Word.Range direct = document.Range(cursor, cursor + text.Length);
                // Never widen this span to include invisible annotation reference CPs:
                // their independent font can make a correct text range report mixed font.
                if (string.Equals(direct.Text, text, StringComparison.Ordinal)) return direct;
            }
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
                find.MatchFuzzy = false;
                find.IgnoreSpace = false;
                find.IgnorePunct = false;
                // Supply the common Find options in one COM call. Keep every literal
                // matching constraint and the final exact-text/gap checks unchanged.
                if (!find.Execute(FindText: chunk.Replace("^", "^^"), Forward: true,
                    Wrap: Word.WdFindWrap.wdFindStop, Format: false, MatchCase: true,
                    MatchWholeWord: false, MatchWildcards: false, MatchSoundsLike: false,
                    MatchAllWordForms: false) || !string.Equals(target.Text, chunk, StringComparison.Ordinal)
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
            if (IsObjectWidth(patch.RuleId)) { ValidateObjectWidth(range, patch); return; }
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
                actual = WordParagraphValueMapper.Alignment((int)range.ParagraphFormat.Alignment);
                EnsureBeforeValue(patch, actual);
                return;
            }
            if (string.Equals(patch.RuleId, "demo/caption-line-spacing", StringComparison.Ordinal))
            {
                if (!string.Equals(patch.DesiredAfter, "single", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("题注行距目标值无效。");
                actual = WordParagraphValueMapper.LineSpacing((int)range.ParagraphFormat.LineSpacingRule,
                    range.ParagraphFormat.LineSpacing);
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
                    ? Math.Abs(actualPoints - expectedPoints) <= (IsObjectWidth(patch.RuleId) ? 0.1 : 0.01)
                    : string.Equals(actual, patch.ExpectedBefore, StringComparison.Ordinal);
            if (!matches)
                throw new InvalidDataException("修复前置条件已变化，已停止本次批量操作：期望「" +
                    patch.ExpectedBefore + "」，当前为「" + actual + "」。");
        }

        private static void ApplyPatch(Word.Range range, EnginePatch patch, Action markMutation)
        {
            if (IsObjectWidth(patch.RuleId)) { ResizeObject(range, patch, markMutation); return; }
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
