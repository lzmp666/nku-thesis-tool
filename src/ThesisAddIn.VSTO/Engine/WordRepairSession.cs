using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using System.Threading.Tasks;
using ThesisAddIn.Protocol;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO.Engine
{
    public static partial class WordDocumentSynchronizer
    {
        // Anchor each grouped issue to the source text with a native Word comment.
        // Word's Resolve action is routed to the add-in repair flow; undo remains an
        // explicit add-in command because Word does not expose custom comment actions.
        private static readonly bool CreateNativeComments = true;

        public sealed class RepairRecord
        {
            internal readonly List<FormatChange> Changes = new List<FormatChange>();
            public int Count => Changes.Count;
            internal void Append(RepairRecord part) { Changes.AddRange(part.Changes); }
        }
        internal sealed class FormatChange
        {
            internal Word.Range Range;
            internal Word.Font Font;
            internal Word.ParagraphFormat Paragraph;
            internal string RuleId;
            internal string Text;
            internal string AppliedSignature;
            internal string FastString;
            internal string FastSecondaryString;
            internal float FastNumber;
            internal int FastEnum;
            internal Action<Action> RestoreObject;
            internal Func<Tuple<string, double[]>> ObjectSignature;
            internal Tuple<string, double[]> AppliedObjectSignature;
            internal Word.Range ObjectTarget;
            internal double ObjectBeforeWidth;
            internal Action<EnginePatch,Action> ApplyObject;
            internal Func<EnginePatch,double[]> DesiredObjectDimensions;
        }
        public static bool ClearPreviousToolMarks(Word.Document document, string snapshot, out string error)
        {
            error = null;
            if (ReadOwnedAnnotations(document).Count == 0 && ReadOwnedHighlightBookmarks(document).Count == 0)
                return true;
            // Never touch annotations or highlights without our recovery metadata.
            return RefreshAnnotations(document, snapshot, new EngineAnalysisResult(), out error);
        }
        public static bool ApplyDirect(Word.Document document, string snapshot,
            IReadOnlyList<EnginePatch> patches, out RepairRecord record, out string error)
        {
            return ApplyDirectWithRanges(document, snapshot, patches, null, out record, out error);
        }

        internal static Dictionary<string, Word.Range> GetRepairRanges(Word.Document document, string snapshot,
            IEnumerable<EnginePatch> patches, Dictionary<string, Word.Range> cached = null) =>
            BuildRanges(document, snapshot, patches.Select(p => p.Location), cached);

        internal static Dictionary<string, Word.Range> GetCachedRepairRanges(Word.Document document,
            IEnumerable<EnginePatch> patches, Dictionary<string, Word.Range> cached,
            Dictionary<string, string> expectedTexts)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (patches == null) throw new ArgumentNullException(nameof(patches));
            if (cached == null || expectedTexts == null)
                throw new InvalidDataException("悬停定位缓存已失效，请重新分析当前文档。");

            var ranges = new Dictionary<string, Word.Range>(StringComparer.Ordinal);
            foreach (EnginePatch patch in patches.Where(value => value != null)
                .GroupBy(value => value.Location, StringComparer.Ordinal).Select(group => group.First()))
            {
                if (string.IsNullOrWhiteSpace(patch.Location) ||
                    !cached.TryGetValue(patch.Location, out Word.Range range) || range == null ||
                    !expectedTexts.TryGetValue(patch.Location, out string expectedText))
                    throw new InvalidDataException("悬停定位缓存缺少目标文本，请重新分析当前文档。");
                try
                {
                    if (range.Document != document || range.Text != expectedText ||
                        range.End - range.Start != expectedText.Length)
                        throw new InvalidDataException("原文已变化，悬停预览已停止；请重新分析。");
                }
                catch (System.Runtime.InteropServices.COMException exception)
                {
                    throw new InvalidDataException("悬停定位缓存已失效，请重新分析当前文档。", exception);
                }
                ranges.Add(patch.Location, range);
            }
            return ranges;
        }

        internal static async Task<Tuple<bool,string>> RemoveResolvedAnnotationsAsync(Word.Document document, IEnumerable<EngineDiagnostic> solved,
            IEnumerable<EngineDiagnostic> remaining, Dictionary<string, Word.Range> verifiedRanges)
        {
            if (!CreateNativeComments)
                return await RemoveResolvedHighlightMarksAsync(document, solved, remaining, verifiedRanges);
            var selected = solved.ToArray();
            var stillActive = new HashSet<string>(remaining.SelectMany(GetOccurrences), StringComparer.Ordinal);
            List<AnnotationRecord> owned = null;
            var remove = new List<AnnotationRecord>();
            try {
                var selectedFields = selected.Select(issue => "\r规则：" + issue.RuleId + "\r期望值：" + issue.Expected + "\r实际值：").ToArray();
                owned = ReadOwnedAnnotations(document, text => selectedFields.Any(field => NormalizeCommentText(text).Contains(field)));
                foreach (EngineDiagnostic issue in selected) {
                    var locations = new HashSet<string>(GetOccurrences(issue), StringComparer.Ordinal);
                    string field = "\r规则：" + issue.RuleId + "\r期望值：" + issue.Expected + "\r实际值：";
                    var matches = owned.Where(record => NormalizeCommentText(record.Text).Contains(field) &&
                        locations.SetEquals(record.OriginalHighlights.Keys)).ToArray();
                    if (matches.Length != 1) throw new InvalidDataException("已修复项的工具批注已被改动，请重新分析。");
                    var match = matches[0];
                    Word.Range anchor = verifiedRanges[GetOccurrences(issue).First()];
                    if (match.Scope.Start != anchor.Start || match.Scope.End != anchor.End)
                        throw new InvalidDataException("已修复项批注的原文锚点已变化。");
                    remove.Add(match);
                }
                string bodyText = document.Content.Text;
                var handled = new HashSet<string>(StringComparer.Ordinal);
                await CooperativeRepairRunner.RunAsync(remove.OrderByDescending(record => record.Index).ToArray(), () => {
                    if (document.Content.Text != bodyText) throw new InvalidDataException("清理批注间隙原文被编辑，转为重新分析。");
                }, record => {
                    string error;
                    bool ok = ExecuteUndoable(document, "清除已修复项的工具批注", () => {
                        Word.Range current = record.Comment.Range.Duplicate;
                        current.TextRetrievalMode.IncludeHiddenText = true;
                        if (NormalizeCommentText(current.Text) != NormalizeCommentText(record.Text))
                            throw new InvalidDataException("工具批注已被编辑，保留改动并重新分析。");
                    }, mark => {
                    foreach (var item in record.OriginalHighlights) {
                        if (stillActive.Contains(item.Key) || !handled.Add(item.Key)) continue;
                        Word.Range range = verifiedRanges[item.Key];
                        int original = (int)HighlightValue(item.Value);
                        string name = HighlightBookmarkName(item.Key, original);
                        if (document.Bookmarks.Exists(name)) {
                            Word.Bookmark bookmark = document.Bookmarks[name];
                            Word.Range stored = bookmark.Range;
                            // Adding a native comment can extend the recovery bookmark
                            // over invisible reference CPs. Formatting ranges must exclude
                            // those CPs, but cleanup may accept a text-identical envelope.
                            if (stored.Start > range.Start || stored.End < range.End || stored.Text != range.Text ||
                                !string.IsNullOrEmpty(document.Range(stored.Start, range.Start).Text) ||
                                !string.IsNullOrEmpty(document.Range(range.End, stored.End).Text))
                                throw new InvalidDataException("原文恢复标记的范围已变化。");
                            // Preserve a highlight the author changed during the repair.
                            if (range.HighlightColorIndex == Word.WdColorIndex.wdYellow) {
                                range.HighlightColorIndex = (Word.WdColorIndex)original; mark();
                            }
                            bookmark.Delete(); mark();
                        } else throw new InvalidDataException("原文恢复标记缺失。");
                    }
                    record.Comment.Delete(); mark();
                    }, out error, suspendRepaint:false);
                    if (!ok) throw new InvalidDataException(error);
                }, () => Task.Delay(1));
                return Tuple.Create(true, (string)null);
            } catch (Exception exception) { return Tuple.Create(false, exception.Message); }
        }

        private static Task<Tuple<bool,string>> RemoveResolvedHighlightMarksAsync(Word.Document document,
            IEnumerable<EngineDiagnostic> solved, IEnumerable<EngineDiagnostic> remaining,
            Dictionary<string, Word.Range> verifiedRanges)
        {
            try
            {
                var stillActive = new HashSet<string>((remaining ?? Enumerable.Empty<EngineDiagnostic>())
                    .SelectMany(GetOccurrences), StringComparer.Ordinal);
                var selectedLocations = (solved ?? Enumerable.Empty<EngineDiagnostic>())
                    .SelectMany(GetOccurrences).Distinct(StringComparer.Ordinal)
                    .Where(location => !stillActive.Contains(location)).ToArray();
                var marks = ReadOwnedHighlightBookmarks(document);
                var cleanup = new List<Tuple<string, HighlightBookmarkRecord, Word.Range>>();
                foreach (string location in selectedLocations)
                {
                    if (verifiedRanges == null || !verifiedRanges.TryGetValue(location, out Word.Range target))
                        throw new InvalidDataException("已修复项的位置缓存已失效，请重新分析。");
                    HighlightBookmarkRecord[] matches = marks.Where(mark => mark.Range != null &&
                        mark.Range.Start <= target.Start && mark.Range.End >= target.End && mark.Range.Text == target.Text &&
                        string.IsNullOrEmpty(document.Range(mark.Range.Start, target.Start).Text) &&
                        string.IsNullOrEmpty(document.Range(target.End, mark.Range.End).Text)).ToArray();
                    if (matches.Length != 1) throw new InvalidDataException("已修复项的原文恢复标记缺失或不唯一，请重新分析。");
                    cleanup.Add(Tuple.Create(location, matches[0], target));
                }
                if (cleanup.Count == 0) return Task.FromResult(Tuple.Create(true, (string)null));

                string originalText = document.Content.Text;
                string error;
                bool ok = ExecuteUndoable(document, "清除已修复项的格式标记", () =>
                {
                    if (document.Content.Text != originalText)
                        throw new InvalidDataException("清理标记期间原文发生变化，已停止操作。");
                    foreach (var item in cleanup)
                    {
                        string name = HighlightBookmarkName(item.Item1, item.Item2.OriginalHighlight);
                        if (!document.Bookmarks.Exists(name) || item.Item3.Text != item.Item2.Range.Text)
                            throw new InvalidDataException("原文恢复标记已变化，已停止清理。");
                    }
                }, markMutation =>
                {
                    foreach (var item in cleanup)
                    {
                        Word.Range target = item.Item3;
                        // Keep an author's later highlight change; restore only our yellow mark.
                        if (target.HighlightColorIndex == Word.WdColorIndex.wdYellow)
                        {
                            target.HighlightColorIndex = (Word.WdColorIndex)item.Item2.OriginalHighlight;
                            markMutation();
                        }
                        string name = HighlightBookmarkName(item.Item1, item.Item2.OriginalHighlight);
                        document.Bookmarks[name].Delete();
                        markMutation();
                    }
                }, out error, suspendRepaint: false);
                return Task.FromResult(Tuple.Create(ok, error));
            }
            catch (Exception exception) { return Task.FromResult(Tuple.Create(false, exception.Message)); }
        }

        internal static bool ApplyDirectWithRanges(Word.Document document, string snapshot,
            IReadOnlyList<EnginePatch> patches, Dictionary<string, Word.Range> verifiedRanges,
            out RepairRecord record, out string error, string actionName = null, bool markDocumentDirty = true,
            bool fastObjectSignature = false, bool suspendRepaint = true, bool fastFormatSnapshot = false)
        {
            var captured = new RepairRecord();
            Dictionary<string, Word.Range> ranges = null;
            bool success = ExecuteUndoable(document, actionName ?? "解决选中的格式问题", () =>
            {
                if (patches == null || patches.Count == 0) throw new InvalidDataException("没有选中可修复的格式。");
                ranges = verifiedRanges ?? BuildRanges(document, snapshot, patches.Select(p => p.Location));
                foreach (EnginePatch patch in patches)
                {
                    Word.Range range = ranges[patch.Location];
                    if (!IsObjectWidth(patch.RuleId)) ValidatePatch(range, patch);
                    // Duplicating both complete COM format objects for every property
                    // needlessly makes Word resolve paragraph and character formatting.
                    bool paragraphProperty = patch.RuleId == "demo/caption-alignment" ||
                        patch.RuleId == "demo/caption-line-spacing" || patch.RuleId == "demo/caption-space-before" ||
                        patch.RuleId == "demo/caption-space-after";
                    var change = new FormatChange { Range = range.Duplicate,
                        Font = !fastFormatSnapshot && !paragraphProperty && !IsObjectWidth(patch.RuleId) ? range.Font.Duplicate : null,
                        Paragraph = !fastFormatSnapshot && paragraphProperty ? range.ParagraphFormat.Duplicate : null,
                        RuleId = patch.RuleId, Text = range.Text };
                    if (fastFormatSnapshot) CaptureFastFormat(change, range);
                    CaptureObjectChange(change, fastObjectSignature);
                    if (IsObjectWidth(patch.RuleId)) ValidateObjectWidth(range,patch,change.ObjectTarget,change.ObjectBeforeWidth);
                    captured.Changes.Add(change);
                }
            }, markMutation =>
            {
                for (int index = 0; index < patches.Count; index++)
                {
                    EnginePatch patch = patches[index];
                    FormatChange change = captured.Changes[index];
                    if (change.ApplyObject != null) change.ApplyObject(patch,markMutation);
                    else ApplyPatch(ranges[patch.Location], patch, markMutation);
                    markMutation();
                    if (change.ObjectSignature != null) {
                        change.AppliedObjectSignature = change.ObjectSignature();
                        if (!WordTableGeometry.EquivalentDimensions(change.DesiredObjectDimensions(patch),change.AppliedObjectSignature.Item2))
                            throw new InvalidDataException("图表写入后回读的尺寸或列宽比例与目标不一致，已撤销本次修改。");
                    }
                    else {
                        ValidatePatch(ranges[patch.Location], new EnginePatch { RuleId = patch.RuleId,
                            ExpectedBefore = patch.DesiredAfter, DesiredAfter = patch.DesiredAfter });
                        change.AppliedSignature = FormatSignature(change);
                    }
                }
                if (markDocumentDirty) document.Saved = false;
            }, out error, suspendRepaint:suspendRepaint);
            record = success ? captured : null;
            return success;
        }
        public static bool UndoRepair(Word.Document document, RepairRecord record, out string error)
        {
            return ExecuteUndoable(document, "撤销该项格式修复", () =>
            {
                if (record == null || record.Count == 0) throw new InvalidDataException("该项没有可撤销的修复记录。");
                foreach (FormatChange change in record.Changes)
                {
                    if (change.Range.Document != document || change.Range.Text != change.Text)
                        throw new InvalidDataException("该处文字或定位范围已被再次编辑，已停止撤销以保留后续改动。请重新分析。");
                    bool unchanged;
                    if (change.ObjectSignature != null)
                    {
                        var current = change.ObjectSignature();
                        unchanged = change.AppliedObjectSignature != null && current.Item1 == change.AppliedObjectSignature.Item1 &&
                            WordTableGeometry.EquivalentDimensions(change.AppliedObjectSignature.Item2, current.Item2);
                    }
                    else unchanged = FormatSignature(change) == change.AppliedSignature;
                    if (!unchanged) throw new InvalidDataException("该处目标格式已被再次编辑，已停止撤销以保留后续改动。请重新分析。");
                }
            }, markMutation =>
            {
                foreach (FormatChange change in record.Changes)
                {
                    RestoreFormat(change, markMutation);
                    markMutation();
                }
                document.Saved = false;
            }, out error, suspendRepaint:true);
        }

        internal static bool UndoHoverPreview(Word.Document document, RepairRecord record,
            bool wasSaved, out string error)
        {
            error = null;
            if (document == null || record == null || record.Count == 0)
            { error = "悬停预览记录已失效。"; return false; }
            try
            {
                // The suggestion list takes keyboard focus while the preview is shown.
                // Check only the affected ranges: reading document.Content.Text here
                // transfers the entire thesis through COM on every pointer movement.
                bool targetUnchanged = record.Changes.All(change => change.Range != null &&
                    change.Range.Document == document && change.Range.Text == change.Text &&
                    IsPreviewPatchUnchanged(change));
                if (targetUnchanged)
                {
                    // The pointer must leave the suggestion list before document input
                    // can occur; this preview custom record is therefore the top undo.
                    // Popping it avoids accumulating undo entries while hovering.
                    bool? screenUpdating = null;
                    try
                    {
                        screenUpdating = document.Application.ScreenUpdating;
                        if (screenUpdating.Value) document.Application.ScreenUpdating = false;
                        dynamic commandBars = document.Application.CommandBars;
                        commandBars.ExecuteMso("Undo");
                        document.Saved = wasSaved;
                        return true;
                    }
                    finally
                    {
                        if (screenUpdating.HasValue)
                        {
                            try { document.Application.ScreenUpdating = screenUpdating.Value; }
                            catch (System.Runtime.InteropServices.COMException) { }
                        }
                    }
                }
                // If the document was edited while the pointer remained over a card,
                // never pop a potentially user-authored undo. Restore only the format
                // range, with UndoRepair's text/format guards.
                if (!UndoRepair(document, record, out error)) return false;
                return true;
            }
            catch (Exception exception)
            {
                error = "无法安全结束悬停预览：" + exception.Message;
                return false;
            }
        }
        private static string FormatSignature(FormatChange change)
        {
            Word.Range range = change.Range;
            switch (change.RuleId)
            {
                case "demo/east-asian-font": case "demo/caption-east-asian-font": return range.Font.NameFarEast;
                case "demo/western-font": return range.Font.NameAscii + "|" + range.Font.NameOther;
                case "demo/font-size": case "demo/caption-font-size": return range.Font.Size.ToString(CultureInfo.InvariantCulture);
                case "demo/caption-alignment": return ((int)range.ParagraphFormat.Alignment).ToString(CultureInfo.InvariantCulture);
                case "demo/caption-line-spacing": return ((int)range.ParagraphFormat.LineSpacingRule).ToString(CultureInfo.InvariantCulture)
                    + "|" + range.ParagraphFormat.LineSpacing.ToString(CultureInfo.InvariantCulture);
                case "demo/caption-space-before": return range.ParagraphFormat.SpaceBefore.ToString(CultureInfo.InvariantCulture);
                case "demo/caption-space-after": return range.ParagraphFormat.SpaceAfter.ToString(CultureInfo.InvariantCulture);
                default: throw new NotSupportedException("该格式尚不支持逐项撤销。");
            }
        }
        private static void RestoreFormat(FormatChange change, Action markMutation)
        {
            if (change.RestoreObject != null) { change.RestoreObject(markMutation); return; }
            Word.Range range = change.Range;
            switch (change.RuleId)
            {
                case "demo/east-asian-font": case "demo/caption-east-asian-font":
                    range.Font.NameFarEast = change.Font != null ? change.Font.NameFarEast : change.FastString; break;
                case "demo/western-font":
                    range.Font.NameAscii = change.Font != null ? change.Font.NameAscii : change.FastString;
                    range.Font.NameOther = change.Font != null ? change.Font.NameOther : change.FastSecondaryString; break;
                case "demo/font-size": case "demo/caption-font-size":
                    range.Font.Size = change.Font != null ? change.Font.Size : change.FastNumber; break;
                case "demo/caption-alignment":
                    range.ParagraphFormat.Alignment = change.Paragraph != null ? change.Paragraph.Alignment :
                        (Word.WdParagraphAlignment)change.FastEnum; break;
                case "demo/caption-line-spacing":
                    range.ParagraphFormat.LineSpacingRule = change.Paragraph != null ? change.Paragraph.LineSpacingRule :
                        (Word.WdLineSpacing)change.FastEnum;
                    range.ParagraphFormat.LineSpacing = change.Paragraph != null ? change.Paragraph.LineSpacing : change.FastNumber; break;
                case "demo/caption-space-before":
                    range.ParagraphFormat.SpaceBefore = change.Paragraph != null ? change.Paragraph.SpaceBefore : change.FastNumber; break;
                case "demo/caption-space-after":
                    range.ParagraphFormat.SpaceAfter = change.Paragraph != null ? change.Paragraph.SpaceAfter : change.FastNumber; break;
                default: throw new NotSupportedException("不能撤销未支持的格式。");
            }
        }

        private static void CaptureFastFormat(FormatChange change, Word.Range range)
        {
            switch (change.RuleId)
            {
                case "demo/east-asian-font": case "demo/caption-east-asian-font":
                    change.FastString = range.Font.NameFarEast; break;
                case "demo/western-font":
                    change.FastString = range.Font.NameAscii;
                    change.FastSecondaryString = range.Font.NameOther; break;
                case "demo/font-size": case "demo/caption-font-size":
                    change.FastNumber = range.Font.Size; break;
                case "demo/caption-alignment":
                    change.FastEnum = (int)range.ParagraphFormat.Alignment; break;
                case "demo/caption-line-spacing":
                    change.FastEnum = (int)range.ParagraphFormat.LineSpacingRule;
                    change.FastNumber = range.ParagraphFormat.LineSpacing; break;
                case "demo/caption-space-before":
                    change.FastNumber = range.ParagraphFormat.SpaceBefore; break;
                case "demo/caption-space-after":
                    change.FastNumber = range.ParagraphFormat.SpaceAfter; break;
                default: throw new NotSupportedException("该格式不支持快速悬停预览。");
            }
        }

        private static bool IsPreviewPatchUnchanged(FormatChange change)
        {
            if (change.ObjectSignature == null)
                return FormatSignature(change) == change.AppliedSignature;
            var current = change.ObjectSignature();
            return change.AppliedObjectSignature != null && current.Item1 == change.AppliedObjectSignature.Item1 &&
                WordTableGeometry.EquivalentDimensions(change.AppliedObjectSignature.Item2, current.Item2);
        }

        private static bool IsObjectWidth(string ruleId) => ruleId == "demo/picture-fit-width" || ruleId == "demo/table-fit-width";

        private static Word.Range AdjacentObjectRange(Word.Range anchor, bool picture)
        {
            Word.Paragraph paragraph = anchor.Paragraphs[1];
            string observed = "";
            for (int index = 0; index < 32; index++)
            {
                paragraph = picture ? paragraph.Previous() : paragraph.Next();
                if (paragraph == null) break;
                Word.Range candidate = paragraph.Range;
                observed = "（行内对象 " + candidate.InlineShapes.Count + "，表格 " + candidate.Tables.Count + "）";
                if (picture && candidate.InlineShapes.Count == 1 && candidate.Tables.Count == 0)
                {
                    // Word's object placeholder is not a portable text character.
                    // Inspect the ranges around the actual object instead.
                    Word.Range shape = candidate.InlineShapes[1].Range;
                    Word.Range before = candidate.Duplicate, after = candidate.Duplicate;
                    before.End = shape.Start; after.Start = shape.End;
                    if (string.IsNullOrWhiteSpace(before.Text) && string.IsNullOrWhiteSpace(after.Text)) return candidate;
                }
                if (!picture && candidate.Tables.Count == 1 && candidate.Tables[1].NestingLevel == 1)
                    return candidate;
                if (candidate.InlineShapes.Count != 0 || candidate.Tables.Count != 0 ||
                    !string.IsNullOrWhiteSpace((candidate.Text ?? "").TrimEnd('\r'))) break;
            }
            throw new InvalidDataException("题注旁没有唯一、可安全定位的图片或简单表格，已停止修改。" + observed);
        }

        private static double ObjectWidth(Word.Range anchor, string ruleId)
        {
            Word.Range target = AdjacentObjectRange(anchor, ruleId == "demo/picture-fit-width");
            if (ruleId == "demo/picture-fit-width") return target.InlineShapes[1].Width;
            Word.Table table = target.Tables[1];
            return ReadTableWidths(table).Sum(width => (double)width);
        }

        private static float[] ReadTableWidths(Word.Table table)
        {
            if (table.AllowAutoFit) throw new InvalidDataException("表格已启用自动调整列宽，已停止缩放。");
            if (table.Rows.Count == 0 || table.Range.Cells.Count > 2000)
                throw new InvalidDataException("表格为空或过大，不能安全原位缩放。");
            var rows = new List<float[]>();
            foreach (Word.Row row in table.Rows)
                rows.Add(row.Cells.Cast<Word.Cell>().Select(cell => cell.Width).ToArray());
            return WordTableGeometry.ReadColumns(rows);
        }

        private static void WriteTableWidths(Word.Table table, float[] widths, Action markMutation)
        {
            // ReadTableWidths has already rejected merged/nonuniform columns. Set
            // each COLUMN once, rather than forcing one layout per cell in every row.
            for (int index = 0; index < widths.Length; index++)
            { table.Columns[index + 1].SetWidth(widths[index], Word.WdRulerStyle.wdAdjustNone); markMutation(); }
        }

        private static void ValidateObjectWidth(Word.Range anchor, EnginePatch patch, Word.Range knownTarget = null, double? knownWidth = null)
        {
            if (!TryParsePoints(patch.DesiredAfter, out double limit)) throw new InvalidDataException("对象宽度目标无效。");
            Word.Range target = knownTarget ?? AdjacentObjectRange(anchor, patch.RuleId == "demo/picture-fit-width");
            Word.PageSetup page = target.Sections[1].PageSetup;
            if (page.TextColumns.Count != 1 || page.Gutter != 0) throw new InvalidDataException("分栏或装订线布局不支持安全缩放。");
            double available = page.PageWidth - page.LeftMargin - page.RightMargin;
            if (patch.RuleId == "demo/picture-fit-width")
                available -= target.ParagraphFormat.LeftIndent + target.ParagraphFormat.RightIndent + target.ParagraphFormat.FirstLineIndent;
            if (Math.Abs(available - limit) > 0.1) throw new InvalidDataException("当前 Word 版面宽度与快照不同，已停止缩放。");
            EnsureBeforeValue(patch, (knownWidth ?? ObjectWidth(anchor, patch.RuleId)).ToString("0.###", CultureInfo.InvariantCulture) + "pt", true);
        }

        private static void ResizeObject(Word.Range anchor, EnginePatch patch, Action markMutation)
        {
            if (!TryParsePoints(patch.DesiredAfter, out double limit)) throw new InvalidDataException("对象宽度目标无效。");
            Word.Range target = AdjacentObjectRange(anchor, patch.RuleId == "demo/picture-fit-width");
            if (patch.RuleId == "demo/picture-fit-width")
            {
                Word.InlineShape picture = target.InlineShapes[1];
                float width = picture.Width, height = picture.Height;
                var locked = picture.LockAspectRatio;
                try
                {
                    if (locked != Microsoft.Office.Core.MsoTriState.msoFalse)
                    { picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse; markMutation(); }
                    picture.Width = (float)limit; markMutation();
                    picture.Height = (float)(height * limit / width); markMutation();
                }
                finally { picture.LockAspectRatio = locked; }
            }
            else
            {
                Word.Table table = target.Tables[1];
                float[] widths = ReadTableWidths(table);
                WriteTableWidths(table, WordTableGeometry.ScaleColumns(widths, limit), markMutation);
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints; markMutation();
                table.PreferredWidth = (float)limit; markMutation();
            }
        }

        private static void CaptureObjectChange(FormatChange change, bool fastSignature = false)
        {
            if (!IsObjectWidth(change.RuleId)) return;
            Word.Range target = AdjacentObjectRange(change.Range, change.RuleId == "demo/picture-fit-width");
            change.ObjectTarget = target;
            if (change.RuleId == "demo/picture-fit-width")
            {
                Word.InlineShape picture = target.InlineShapes[1];
                float width = picture.Width, height = picture.Height;
                var locked = picture.LockAspectRatio;
                change.ObjectBeforeWidth = width;
                change.DesiredObjectDimensions = patch => {
                    double limit; if (!TryParsePoints(patch.DesiredAfter,out limit)) throw new InvalidDataException("图片目标宽度无效。");
                    return new[]{limit,height*limit/width};
                };
                change.ApplyObject = (patch,mark) => {
                    double[] dimensions = change.DesiredObjectDimensions(patch);
                    try {
                        if (locked != Microsoft.Office.Core.MsoTriState.msoFalse) { picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse; mark(); }
                        picture.Width = (float)dimensions[0]; mark(); picture.Height = (float)dimensions[1]; mark();
                    } finally { picture.LockAspectRatio = locked; }
                };
                change.ObjectSignature = fastSignature
                    ? (Func<Tuple<string, double[]>>)(() => Tuple.Create(picture.LockAspectRatio.ToString(),
                        new[] { (double)picture.Width, picture.Height }))
                    : () => Tuple.Create(picture.LockAspectRatio + "|" + PictureContentSignature(picture),
                        new[] { (double)picture.Width, picture.Height });
                change.RestoreObject = markMutation =>
                {
                    try
                    {
                        if (picture.LockAspectRatio != Microsoft.Office.Core.MsoTriState.msoFalse)
                        { picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse; markMutation(); }
                        picture.Width = width; markMutation(); picture.Height = height; markMutation();
                    }
                    finally { picture.LockAspectRatio = locked; }
                };
            }
            else
            {
                Word.Table table = target.Tables[1];
                float[] widths = ReadTableWidths(table);
                var preferredType = table.PreferredWidthType;
                float preferredWidth = table.PreferredWidth;
                change.ObjectBeforeWidth = widths.Sum(value => (double)value);
                change.DesiredObjectDimensions = patch => {
                    double limit; if (!TryParsePoints(patch.DesiredAfter,out limit)) throw new InvalidDataException("表格目标宽度无效。");
                    return new[]{limit}.Concat(WordTableGeometry.ScaleColumns(widths,limit).Select(value => (double)value)).ToArray();
                };
                change.ApplyObject = (patch,mark) => {
                    double[] dimensions = change.DesiredObjectDimensions(patch);
                    WriteTableWidths(table,dimensions.Skip(1).Select(value => (float)value).ToArray(),mark);
                    table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints; mark();
                    table.PreferredWidth = (float)dimensions[0]; mark();
                };
                change.ObjectSignature = fastSignature
                    ? (Func<Tuple<string, double[]>>)(() => Tuple.Create(table.Rows.Count + "|" + table.AllowAutoFit + "|" +
                        table.PreferredWidthType, new[] { (double)table.PreferredWidth }.Concat(
                            ReadTableWidths(table).Select(value => (double)value)).ToArray()))
                    : () => Tuple.Create(table.Range.Text + "|" + table.Rows.Count + "|" + table.AllowAutoFit + "|" +
                        table.PreferredWidthType, new[] { (double)table.PreferredWidth }.Concat(
                            ReadTableWidths(table).Select(value => (double)value)).ToArray());
                change.RestoreObject = markMutation =>
                {
                    WriteTableWidths(table, widths, markMutation);
                    table.PreferredWidthType = preferredType; markMutation(); table.PreferredWidth = preferredWidth; markMutation();
                };
            }
        }

        private static string PictureContentSignature(Word.InlineShape picture)
        {
            // Ignore unrelated Flat OPC properties/styles/statistics. A changed image
            // must block item undo; another format repair must not block it.
            XNamespace package = "http://schemas.microsoft.com/office/2006/xmlPackage";
            // A range consisting solely of the object marker can omit media
            // payloads. Export its image-only paragraph to include dependencies.
            XDocument xml = XDocument.Parse(picture.Range.Paragraphs[1].Range.WordOpenXML);
            string media = string.Join("|", xml.Descendants(package + "part")
                .Where(part => ((string)part.Attribute(package + "contentType") ?? "").StartsWith("image/", StringComparison.Ordinal))
                .Select(part => part.Element(package + "binaryData")?.Value ?? "")
                .Concat(xml.Descendants(XName.Get("binData", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).Select(element => element.Value))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => Convert.ToBase64String(Convert.FromBase64String(value)))
                .OrderBy(value => value, StringComparer.Ordinal));
            if (string.IsNullOrWhiteSpace(media)) throw new InvalidDataException("无法核对图片内容，已停止撤销。图片段落部件数：" + xml.Descendants(package + "part").Count());
            using (SHA256 hash = SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(media)));
        }
    }
}
