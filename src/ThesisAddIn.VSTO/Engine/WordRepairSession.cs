using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ThesisAddIn.Protocol;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO.Engine
{
    public static partial class WordDocumentSynchronizer
    {
        // Actionable cards are hosted by the add-in, not Word's reply/resolve UI.
        private static readonly bool CreateNativeComments = false;

        public sealed class RepairRecord
        {
            internal readonly List<FormatChange> Changes = new List<FormatChange>();
            public int Count => Changes.Count;
        }
        internal sealed class FormatChange
        {
            internal Word.Range Range;
            internal Word.Font Font;
            internal Word.ParagraphFormat Paragraph;
            internal string RuleId;
            internal string Text;
            internal string AppliedSignature;
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
            var captured = new RepairRecord();
            Dictionary<string, Word.Range> ranges = null;
            bool success = ExecuteUndoable(document, "解决选中的格式问题", () =>
            {
                if (patches == null || patches.Count == 0) throw new InvalidDataException("没有选中可修复的格式。");
                ranges = BuildRanges(document, snapshot, patches.Select(p => p.Location));
                foreach (EnginePatch patch in patches)
                {
                    Word.Range range = ranges[patch.Location];
                    ValidatePatch(range, patch);
                    captured.Changes.Add(new FormatChange { Range = range.Duplicate, Font = range.Font.Duplicate,
                        Paragraph = range.ParagraphFormat.Duplicate, RuleId = patch.RuleId, Text = range.Text });
                }
            }, markMutation =>
            {
                for (int index = 0; index < patches.Count; index++)
                {
                    EnginePatch patch = patches[index];
                    ApplyPatch(ranges[patch.Location], patch);
                    markMutation();
                    ValidatePatch(ranges[patch.Location], new EnginePatch { RuleId = patch.RuleId,
                        ExpectedBefore = patch.DesiredAfter, DesiredAfter = patch.DesiredAfter });
                    captured.Changes[index].AppliedSignature = FormatSignature(captured.Changes[index]);
                }
                document.Saved = false;
            }, out error);
            record = success ? captured : null;
            return success;
        }
        public static bool UndoRepair(Word.Document document, RepairRecord record, out string error)
        {
            return ExecuteUndoable(document, "撤销该项格式修复", () =>
            {
                if (record == null || record.Count == 0) throw new InvalidDataException("该项没有可撤销的修复记录。");
                foreach (FormatChange change in record.Changes)
                    if (change.Range.Document != document || change.Range.Text != change.Text ||
                        FormatSignature(change) != change.AppliedSignature)
                        throw new InvalidDataException("该处文字或目标格式已被再次编辑，已停止撤销以保留后续改动。请重新分析。");
            }, markMutation =>
            {
                foreach (FormatChange change in record.Changes)
                {
                    RestoreFormat(change);
                    markMutation();
                }
                document.Saved = false;
            }, out error);
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
        private static void RestoreFormat(FormatChange change)
        {
            Word.Range range = change.Range;
            switch (change.RuleId)
            {
                case "demo/east-asian-font": case "demo/caption-east-asian-font": range.Font.NameFarEast = change.Font.NameFarEast; break;
                case "demo/western-font": range.Font.NameAscii = change.Font.NameAscii; range.Font.NameOther = change.Font.NameOther; break;
                case "demo/font-size": case "demo/caption-font-size": range.Font.Size = change.Font.Size; break;
                case "demo/caption-alignment": range.ParagraphFormat.Alignment = change.Paragraph.Alignment; break;
                case "demo/caption-line-spacing": range.ParagraphFormat.LineSpacingRule = change.Paragraph.LineSpacingRule;
                    range.ParagraphFormat.LineSpacing = change.Paragraph.LineSpacing; break;
                case "demo/caption-space-before": range.ParagraphFormat.SpaceBefore = change.Paragraph.SpaceBefore; break;
                case "demo/caption-space-after": range.ParagraphFormat.SpaceAfter = change.Paragraph.SpaceAfter; break;
                default: throw new NotSupportedException("不能撤销未支持的格式。");
            }
        }
    }
}
