using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ThesisAddIn.Protocol;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO.Engine
{
    public static partial class WordDocumentSynchronizer
    {
        // Only the explicitly selected target properties are captured/restored. A
        // later italic/colour change must not be erased by whole-thesis undo.
        internal sealed class OutlineFormattingRecord
        {
            internal readonly List<OutlineParagraphChange> Changes = new List<OutlineParagraphChange>();
            internal int Count => Changes.Count;
        }
        internal sealed class OutlineParagraphChange
        {
            internal EngineStructureFormatOperation Operation;
            internal Word.Range Paragraph;
            internal Word.ParagraphFormat BeforeParagraph;
            internal readonly List<OutlineRunChange> Runs = new List<OutlineRunChange>();
            internal string BeforeSignature, AppliedSignature;
        }
        internal sealed class OutlineRunChange
        {
            internal Word.Range Range;
            internal Word.Font BeforeFont;
            internal string Text;
        }

        internal static void LocateOutlineBlock(Word.Document document, string snapshot,
            EngineStructureSourceBlock block)
        {
            if (block == null || string.IsNullOrEmpty(block.FirstRunLocation))
                throw new InvalidDataException("该块没有可安全定位的文字范围。");
            Word.Range range = BuildRanges(document, snapshot, new[] { block.FirstRunLocation })[block.FirstRunLocation];
            Word.Range paragraph = range.Paragraphs[1].Range.Duplicate;
            if ((paragraph.Text ?? string.Empty).TrimEnd('\r', '\a') != block.ExactText)
                throw new InvalidDataException("大纲位置与当前原文不一致，请重新识别大纲。");
            range.Select();
            document.ActiveWindow.ScrollIntoView(range);
        }

        internal static async Task<OutlineFormattingRecord> ApplyOutlineFormattingAsync(
            Word.Document document, string snapshot, EngineStructureFormatPlan plan,
            Action validateDocument, Action<int, int> progress, OutlineFormattingRecord record, Action validateBeforeWrite)
        {
            if (plan == null || plan.Operations == null || plan.Operations.Count == 0)
                throw new InvalidDataException("确认的大纲没有可应用的排版项目。");
            if (plan.Operations.Select(item => item.BlockId).Distinct(StringComparer.Ordinal).Count() != plan.Operations.Count)
                throw new InvalidDataException("整篇排版计划包含重复原文块。");
            ValidateOutlineDocument(document);
            validateDocument();
            string originalText = document.Content.Text;
            var mapping = PrepareRangeMapping(snapshot);
            var prepared = new List<OutlineParagraphChange>();
            await CooperativeRepairRunner.RunAsync(plan.Operations, () => {
                validateDocument(); ValidateOutlineDocument(document);
                if (document.Content.Text != originalText) throw new InvalidDataException("核对大纲期间原文文字被再次编辑。");
            }, operation => {
                ValidateOutlineTarget(operation.Target);
                if (operation.IsTableCell || !operation.HasCompleteTextMapping || operation.RunLocations.Count == 0)
                    throw new InvalidDataException("整篇排版计划包含未支持的原文结构。");
                var ranges = BuildRanges(document, snapshot, operation.RunLocations, prepared: mapping);
                Word.Range paragraph = ranges[operation.RunLocations[0]].Paragraphs[1].Range.Duplicate;
                Word.Style style = paragraph.get_Style() as Word.Style;
                if (style != null && style.AutomaticallyUpdate)
                    throw new InvalidDataException("所选段落启用了样式自动更新，可能影响未勾选的内容。请先由作者关闭该样式的自动更新，再排版。");
                if ((paragraph.Text ?? string.Empty).TrimEnd('\r', '\a') != operation.ExactText)
                    throw new InvalidDataException("原文块已改变，整篇排版没有开始。");
                var change = new OutlineParagraphChange { Operation = operation, Paragraph = paragraph,
                    BeforeParagraph = paragraph.ParagraphFormat.Duplicate };
                foreach (string location in operation.RunLocations)
                {
                    Word.Range run = ranges[location];
                    if (run.Start < paragraph.Start || run.End > paragraph.End)
                        throw new InvalidDataException("原文字符范围越出段落。");
                    change.Runs.Add(new OutlineRunChange { Range = run.Duplicate, BeforeFont = run.Font.Duplicate, Text = run.Text });
                }
                change.BeforeSignature = OutlineSignature(change);
                prepared.Add(change);
            }, () => Task.Delay(1), count => progress?.Invoke(0, plan.Operations.Count));
            // Preparation yields to Word; revalidate the full original snapshot
            // once before the first mutation, including edits to not-yet-captured styles.
            validateBeforeWrite();
            if (record == null || record.Count != 0) throw new InvalidDataException("整篇排版记录无效。");
            try
            {
                await CooperativeRepairRunner.RunAsync(prepared, () => {
                    validateDocument(); ValidateOutlineDocument(document);
                    if (document.Content.Text != originalText)
                        throw new InvalidDataException("排版期间原文文字发生变化，已停止并恢复本轮已修改格式。");
                }, change => {
                    string error;
                    if (!ExecuteUndoable(document, "整篇排版：" + change.Operation.Section, () => {
                        ValidateOutlineText(change);
                        if (OutlineSignature(change) != change.BeforeSignature)
                            throw new InvalidDataException("待排版段落的格式被再次编辑，已停止以保留后续改动。");
                    }, mark => {
                        WriteOutlineTarget(change, mark);
                        ValidateOutlineText(change);
                        if (!OutlineTargetMatches(change))
                            throw new InvalidDataException("Word 回读格式与目标不一致。");
                        change.AppliedSignature = OutlineSignature(change);
                    }, out error, suspendRepaint: false)) throw new InvalidDataException(error);
                    record.Changes.Add(change);
                }, () => Task.Delay(1), count => progress?.Invoke(count, prepared.Count));
                document.Saved = false;
                return record;
            }
            catch (Exception exception)
            {
                var rollback = record.Count == 0 ? Tuple.Create(true, (string)null) :
                    await UndoOutlineFormattingAsync(document, record, () => { });
                if (!rollback.Item1)
                    throw new InvalidDataException(exception.Message + "；已完成段落未能全部恢复：" + rollback.Item2, exception);
                throw;
            }
        }

        internal static async Task<Tuple<bool, string>> UndoOutlineFormattingAsync(Word.Document document,
            OutlineFormattingRecord record, Action validateDocument)
        {
            try {
                ValidateOutlineDocument(document);
                if (record == null || record.Count == 0) throw new InvalidDataException("当前没有整篇排版撤销记录。");
                var pending = record.Changes.AsEnumerable().Reverse().ToArray();
                await CooperativeRepairRunner.RunAsync(pending, validateDocument, change => {
                    ValidateOutlineText(change);
                    if (OutlineSignature(change) != change.AppliedSignature)
                        throw new InvalidDataException("排版涉及的格式已被再次修改，停止撤销以保留后续编辑。");
                }, () => Task.Delay(1));
                await CooperativeRepairRunner.RunAsync(pending, validateDocument, change => {
                    string error;
                    if (!ExecuteUndoable(document, "撤销整篇排版段落", () => {
                        ValidateOutlineDocument(document); ValidateOutlineText(change);
                        if (OutlineSignature(change) != change.AppliedSignature)
                            throw new InvalidDataException("该段格式被再次编辑，已停止撤销。");
                    }, mark => {
                        RestoreOutlineTarget(change, mark);
                        ValidateOutlineText(change);
                        if (OutlineSignature(change) != change.BeforeSignature)
                            throw new InvalidDataException("第 " + (change.Operation.ParagraphIndex + 1) +
                                " 段恢复后的格式与排版前不一致，已恢复该段排版结果。原值=" + change.BeforeSignature.Replace('\u001f', '|') +
                                "；回读=" + OutlineSignature(change).Replace('\u001f', '|'));
                    }, out error, suspendRepaint: false)) throw new InvalidDataException(error);
                    record.Changes.Remove(change);
                }, () => Task.Delay(1));
                document.Saved = false;
                return Tuple.Create(true, (string)null);
            } catch (Exception exception) { return Tuple.Create(false, exception.Message); }
        }

        private static void ValidateOutlineDocument(Word.Document document)
        {
            if (document == null || document.ReadOnly || document.ProtectionType != Word.WdProtectionType.wdNoProtection)
                throw new InvalidDataException("当前文档只读、受保护或已关闭，无法排版。");
            if (document.TrackRevisions || document.Revisions.Count != 0)
                throw new InvalidDataException("请先由作者处理现有修订并关闭修订模式，再进行整篇排版。");
        }
        private static void ValidateOutlineText(OutlineParagraphChange change)
        {
            if ((change.Paragraph.Text ?? string.Empty).TrimEnd('\r', '\a') != change.Operation.ExactText ||
                change.Runs.Any(run => run.Range.Text != run.Text))
                throw new InvalidDataException("排版涉及的原文已改变，请重新识别大纲。");
        }
        private static void ValidateOutlineTarget(EngineStructureFormatTarget target)
        {
            if (target == null || target.SizePt.HasValue && (target.SizePt < 1 || target.SizePt > 1638) ||
                target.ExactLineSpacingPt.HasValue && (target.ExactLineSpacingPt <= 0 || target.ExactLineSpacingPt > 1584) ||
                target.ExactLineSpacingPt.HasValue && target.SingleLineSpacing == true ||
                target.FirstLineChars.HasValue && Math.Abs(target.FirstLineChars.Value) > 20 ||
                target.LeftIndentChars.HasValue && Math.Abs(target.LeftIndentChars.Value) > 20 ||
                target.Alignment.HasValue && !Enum.IsDefined(typeof(EngineThesisRoleAlignment), target.Alignment.Value))
                throw new InvalidDataException("整篇排版目标无效。");
        }
        private static string OutlineSignature(OutlineParagraphChange change)
        {
            var target = change.Operation.Target;
            var values = new List<string>();
            foreach (var run in change.Runs)
            {
                Word.Font font = run.Range.Font;
                if (target.EastAsiaFont != null) values.Add(font.NameFarEast ?? "<null>");
                if (target.LatinFont != null) { values.Add(font.NameAscii ?? "<null>"); values.Add(font.NameOther ?? "<null>"); }
                if (target.SizePt.HasValue) { values.Add(Number(font.Size)); values.Add(Number(font.SizeBi)); }
                if (target.Bold.HasValue) { values.Add(font.Bold.ToString()); values.Add(font.BoldBi.ToString()); }
            }
            Word.ParagraphFormat p = change.Paragraph.ParagraphFormat;
            if (target.Alignment.HasValue) values.Add(((int)p.Alignment).ToString());
            if (target.ExactLineSpacingPt.HasValue || target.SingleLineSpacing.HasValue) {
                values.Add(((int)p.LineSpacingRule).ToString());
                // Single/1.5/double point heights are derived from font metrics,
                // not an independently restorable paragraph property.
                if (p.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceAtLeast ||
                    p.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceExactly ||
                    p.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceMultiple) values.Add(Number(p.LineSpacing));
            }
            if (target.BeforePt.HasValue) { values.Add(Number(p.SpaceBefore)); values.Add(p.SpaceBeforeAuto.ToString()); }
            if (target.AfterPt.HasValue) { values.Add(Number(p.SpaceAfter)); values.Add(p.SpaceAfterAuto.ToString()); }
            if (target.FirstLineChars.HasValue) { values.Add(Number(p.FirstLineIndent)); values.Add(Number(p.CharacterUnitFirstLineIndent)); }
            if (target.LeftIndentChars.HasValue) { values.Add(Number(p.LeftIndent)); values.Add(Number(p.CharacterUnitLeftIndent)); }
            return string.Join("\u001f", values);
        }
        private static string Number(float value) => value.ToString("R", CultureInfo.InvariantCulture);
        private static bool Close(float actual, double expected) => Math.Abs(actual - expected) <= 0.06;
        private static Word.WdParagraphAlignment OutlineAlignment(EngineThesisRoleAlignment value)
        {
            switch (value) {
                case EngineThesisRoleAlignment.Left: return Word.WdParagraphAlignment.wdAlignParagraphLeft;
                case EngineThesisRoleAlignment.Center: return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case EngineThesisRoleAlignment.Right: return Word.WdParagraphAlignment.wdAlignParagraphRight;
                case EngineThesisRoleAlignment.Justified: return Word.WdParagraphAlignment.wdAlignParagraphJustify;
                default: throw new InvalidDataException("段落对齐目标无效。");
            }
        }
        private static void WriteOutlineTarget(OutlineParagraphChange change, Action mark)
        {
            var target = change.Operation.Target;
            foreach (var run in change.Runs)
            {
                Word.Font font = run.Range.Font;
                if (target.EastAsiaFont != null) { font.NameFarEast = target.EastAsiaFont; mark(); }
                if (target.LatinFont != null) { font.NameAscii = target.LatinFont; mark(); font.NameOther = target.LatinFont; }
                if (target.SizePt.HasValue) { font.Size = (float)target.SizePt.Value; mark(); font.SizeBi = (float)target.SizePt.Value; }
                if (target.Bold.HasValue) { font.Bold = target.Bold.Value ? -1 : 0; mark(); font.BoldBi = target.Bold.Value ? -1 : 0; }
            }
            Word.ParagraphFormat p = change.Paragraph.ParagraphFormat;
            if (target.Alignment.HasValue) { p.Alignment = OutlineAlignment(target.Alignment.Value); mark(); }
            if (target.ExactLineSpacingPt.HasValue) { p.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly; mark(); p.LineSpacing = (float)target.ExactLineSpacingPt.Value; }
            else if (target.SingleLineSpacing == true) { p.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle; mark(); }
            if (target.BeforePt.HasValue) { p.SpaceBeforeAuto = 0; mark(); p.SpaceBefore = (float)target.BeforePt.Value; }
            if (target.AfterPt.HasValue) { p.SpaceAfterAuto = 0; mark(); p.SpaceAfter = (float)target.AfterPt.Value; }
            if (target.FirstLineChars.HasValue) { p.FirstLineIndent = 0; mark(); p.CharacterUnitFirstLineIndent = target.FirstLineChars.Value; }
            if (target.LeftIndentChars.HasValue) { p.LeftIndent = 0; mark(); p.CharacterUnitLeftIndent = target.LeftIndentChars.Value; }
        }
        private static bool OutlineTargetMatches(OutlineParagraphChange change)
        {
            var target = change.Operation.Target;
            foreach (var run in change.Runs)
            {
                Word.Font font = run.Range.Font;
                if (target.EastAsiaFont != null && font.NameFarEast != target.EastAsiaFont ||
                    target.LatinFont != null && (font.NameAscii != target.LatinFont || font.NameOther != target.LatinFont) ||
                    target.SizePt.HasValue && (!Close(font.Size, target.SizePt.Value) || !Close(font.SizeBi, target.SizePt.Value)) ||
                    target.Bold.HasValue && (font.Bold != (target.Bold.Value ? -1 : 0) || font.BoldBi != (target.Bold.Value ? -1 : 0))) return false;
            }
            Word.ParagraphFormat p = change.Paragraph.ParagraphFormat;
            return (!target.Alignment.HasValue || p.Alignment == OutlineAlignment(target.Alignment.Value)) &&
                (!target.ExactLineSpacingPt.HasValue || p.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceExactly && Close(p.LineSpacing, target.ExactLineSpacingPt.Value)) &&
                (target.SingleLineSpacing != true || p.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceSingle) &&
                (!target.BeforePt.HasValue || p.SpaceBeforeAuto == 0 && Close(p.SpaceBefore, target.BeforePt.Value)) &&
                (!target.AfterPt.HasValue || p.SpaceAfterAuto == 0 && Close(p.SpaceAfter, target.AfterPt.Value)) &&
                (!target.FirstLineChars.HasValue || Close(p.CharacterUnitFirstLineIndent, target.FirstLineChars.Value)) &&
                (!target.LeftIndentChars.HasValue || Close(p.CharacterUnitLeftIndent, target.LeftIndentChars.Value));
        }
        private static void RestoreOutlineTarget(OutlineParagraphChange change, Action mark)
        {
            var target = change.Operation.Target;
            foreach (var run in change.Runs)
            {
                Word.Font font = run.Range.Font, before = run.BeforeFont;
                if (target.EastAsiaFont != null) { font.NameFarEast = before.NameFarEast; mark(); }
                if (target.LatinFont != null) { font.NameAscii = before.NameAscii; mark(); font.NameOther = before.NameOther; }
                if (target.SizePt.HasValue) { font.Size = before.Size; mark(); font.SizeBi = before.SizeBi; }
                if (target.Bold.HasValue) { font.Bold = before.Bold; mark(); font.BoldBi = before.BoldBi; }
            }
            Word.ParagraphFormat p = change.Paragraph.ParagraphFormat, original = change.BeforeParagraph;
            if (target.Alignment.HasValue) { p.Alignment = original.Alignment; mark(); }
            if (target.ExactLineSpacingPt.HasValue || target.SingleLineSpacing.HasValue) {
                p.LineSpacingRule = original.LineSpacingRule; mark();
                if (original.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceAtLeast ||
                    original.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceExactly ||
                    original.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceMultiple)
                    p.LineSpacing = original.LineSpacing;
            }
            if (target.BeforePt.HasValue) { p.SpaceBefore = original.SpaceBefore; mark(); p.SpaceBeforeAuto = original.SpaceBeforeAuto; }
            if (target.AfterPt.HasValue) { p.SpaceAfter = original.SpaceAfter; mark(); p.SpaceAfterAuto = original.SpaceAfterAuto; }
            if (target.FirstLineChars.HasValue) {
                p.CharacterUnitFirstLineIndent = 0; mark();
                p.FirstLineIndent = original.FirstLineIndent;
                if (original.CharacterUnitFirstLineIndent != 0) p.CharacterUnitFirstLineIndent = original.CharacterUnitFirstLineIndent;
            }
            if (target.LeftIndentChars.HasValue) {
                p.CharacterUnitLeftIndent = 0; mark();
                p.LeftIndent = original.LeftIndent;
                if (original.CharacterUnitLeftIndent != 0) p.CharacterUnitLeftIndent = original.CharacterUnitLeftIndent;
            }
        }
    }
}
