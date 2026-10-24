using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using ThesisAddIn.Protocol;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO.Engine
{
    public static partial class WordDocumentSynchronizer
    {
        // Actionable cards are hosted by the add-in, not Word's reply/resolve UI.
        private static readonly bool CreateNativeComments = true;

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
            internal Action<Action> RestoreObject;
            internal Func<Tuple<string, double[]>> ObjectSignature;
            internal Tuple<string, double[]> AppliedObjectSignature;
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
                    var change = new FormatChange { Range = range.Duplicate, Font = range.Font.Duplicate,
                        Paragraph = range.ParagraphFormat.Duplicate, RuleId = patch.RuleId, Text = range.Text };
                    CaptureObjectChange(change);
                    captured.Changes.Add(change);
                }
            }, markMutation =>
            {
                for (int index = 0; index < patches.Count; index++)
                {
                    EnginePatch patch = patches[index];
                    ApplyPatch(ranges[patch.Location], patch, markMutation);
                    markMutation();
                    ValidatePatch(ranges[patch.Location], new EnginePatch { RuleId = patch.RuleId,
                        ExpectedBefore = patch.DesiredAfter, DesiredAfter = patch.DesiredAfter });
                    FormatChange change = captured.Changes[index];
                    if (change.ObjectSignature != null) change.AppliedObjectSignature = change.ObjectSignature();
                    else change.AppliedSignature = FormatSignature(change);
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
        private static void RestoreFormat(FormatChange change, Action markMutation)
        {
            if (change.RestoreObject != null) { change.RestoreObject(markMutation); return; }
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
            foreach (Word.Row row in table.Rows)
                for (int index = 0; index < widths.Length; index++)
                { row.Cells[index + 1].SetWidth(widths[index], Word.WdRulerStyle.wdAdjustNone); markMutation(); }
        }

        private static void ValidateObjectWidth(Word.Range anchor, EnginePatch patch)
        {
            if (!TryParsePoints(patch.DesiredAfter, out double limit)) throw new InvalidDataException("对象宽度目标无效。");
            Word.Range target = AdjacentObjectRange(anchor, patch.RuleId == "demo/picture-fit-width");
            Word.PageSetup page = target.Sections[1].PageSetup;
            if (page.TextColumns.Count != 1 || page.Gutter != 0) throw new InvalidDataException("分栏或装订线布局不支持安全缩放。");
            double available = page.PageWidth - page.LeftMargin - page.RightMargin;
            if (patch.RuleId == "demo/picture-fit-width")
                available -= target.ParagraphFormat.LeftIndent + target.ParagraphFormat.RightIndent + target.ParagraphFormat.FirstLineIndent;
            if (Math.Abs(available - limit) > 0.1) throw new InvalidDataException("当前 Word 版面宽度与快照不同，已停止缩放。");
            EnsureBeforeValue(patch, ObjectWidth(anchor, patch.RuleId).ToString("0.###", CultureInfo.InvariantCulture) + "pt", true);
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

        private static void CaptureObjectChange(FormatChange change)
        {
            if (!IsObjectWidth(change.RuleId)) return;
            Word.Range target = AdjacentObjectRange(change.Range, change.RuleId == "demo/picture-fit-width");
            if (change.RuleId == "demo/picture-fit-width")
            {
                Word.InlineShape picture = target.InlineShapes[1];
                float width = picture.Width, height = picture.Height;
                var locked = picture.LockAspectRatio;
                change.ObjectSignature = () => Tuple.Create(picture.LockAspectRatio + "|" + PictureContentSignature(picture),
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
                change.ObjectSignature = () => Tuple.Create(table.Range.Text + "|" + table.Rows.Count + "|" + table.AllowAutoFit + "|" +
                    table.PreferredWidthType, new[] { (double)table.PreferredWidth }.Concat(ReadTableWidths(table).Select(value => (double)value)).ToArray());
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
