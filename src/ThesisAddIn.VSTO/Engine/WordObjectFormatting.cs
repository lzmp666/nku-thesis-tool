using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO.Engine
{
    public static partial class WordDocumentSynchronizer
    {
        private sealed class ObjectWork
        {
            internal Word.InlineShape Picture;
            internal Word.Table Table;
            internal Word.Range Caption;
            internal string Name;
        }

        internal static async Task<string> FormatObjectsAsync(Word.Document document, ObjectFormattingSettings settings,
            Action<string> progress, Action validate, Action<string> trace)
        {
            if (document.ReadOnly || document.ProtectionType != Word.WdProtectionType.wdNoProtection)
                throw new InvalidDataException("当前文档只读或受保护。");
            string text = document.Content.Text;
            var work = new List<ObjectWork>();
            var skipped = new List<string>();
            foreach (Word.InlineShape picture in document.InlineShapes)
            {
                if (picture.Type != Word.WdInlineShapeType.wdInlineShapePicture &&
                    picture.Type != Word.WdInlineShapeType.wdInlineShapeLinkedPicture) {
                    skipped.Add("图形：不是普通行内图片（图表、公式或其他对象）"); continue;
                }
                if (picture.Range.Tables.Count > 0) { skipped.Add("图片：位于表格内"); continue; }
                Word.Range paragraph = picture.Range.Paragraphs[1].Range;
                Word.Range before = paragraph.Duplicate, after = paragraph.Duplicate;
                before.End = picture.Range.Start; after.Start = picture.Range.End;
                Word.Range caption = FindObjectCaption(paragraph.Paragraphs[1], true);
                if (paragraph.InlineShapes.Count != 1 || !string.IsNullOrWhiteSpace(before.Text) ||
                    !string.IsNullOrWhiteSpace(after.Text) || caption == null) {
                    skipped.Add("图片：没有独立段落及下方唯一图题，需人工确认对应关系/题注位置"); continue;
                }
                work.Add(new ObjectWork { Picture = picture, Caption = caption, Name = "图片 " + (caption.Text ?? "").Trim() });
            }
            foreach (Word.Table table in document.Tables)
            {
                if (table.NestingLevel != 1) { skipped.Add("表格：嵌套表格"); continue; }
                Word.Range caption = FindObjectCaption(table.Range.Paragraphs[1], false);
                if (caption == null) { skipped.Add("表格：没有上方唯一表题，需人工确认对应关系/题注位置"); continue; }
                string rejection = ObjectFormattingPolicy.TableRejection(table.Range.WordOpenXML, table.Range.Text);
                if (rejection != null) { skipped.Add((caption.Text ?? "表格").Trim() + "：" + rejection); continue; }
                work.Add(new ObjectWork { Table = table, Caption = caption, Name = "表格 " + (caption.Text ?? "").Trim() });
            }
            int completed = 0;
            await CooperativeRepairRunner.RunAsync(work, () => {
                validate();
                if (document.Content.Text != text) throw new InvalidDataException("排版间隙原文被编辑，已停止后续操作；已完成对象可用 Ctrl+Z 撤销。");
            }, item => {
                var timer = Stopwatch.StartNew();
                string error;
                bool success = ExecuteUndoable(document, "图表排版：" + item.Name, () => {
                    if (item.Table != null && ObjectFormattingPolicy.TableRejection(item.Table.Range.WordOpenXML, item.Table.Range.Text) != null)
                        throw new InvalidDataException("表格结构已变化。");
                }, mark => {
                    string originalText = item.Table?.Range.Text ?? item.Picture.Range.Paragraphs[1].Range.Text;
                    if (item.Picture != null) FormatPicture(item.Picture, settings, mark);
                    else FormatTable(item.Table, settings, mark, skipped);
                    if (settings.Captions) FormatCaption(item.Caption, item.Picture != null, mark);
                    string currentText = item.Table?.Range.Text ?? item.Picture.Range.Paragraphs[1].Range.Text;
                    if (currentText != originalText) throw new InvalidDataException("对象文字发生变化，已撤销该对象排版。");
                    document.Saved = false;
                }, out error);
                trace?.Invoke("图表排版 object=" + item.Name + " write=" + timer.ElapsedMilliseconds + "ms ok=" + success);
                if (!success) throw new InvalidDataException(error + " 已完成 " + completed + " 个对象；可用 Ctrl+Z 逐个撤销。");
                completed++;
                progress?.Invoke("图表排版 " + completed + "/" + work.Count + "：" + item.Name);
            }, () => Task.Delay(1));
            if (document.Shapes.Count > 0) skipped.Add("浮动图片/组合流程图 " + document.Shapes.Count + " 个：保留锚点与环绕，未自动修改");
            return "已排版 " + completed + " 个配对图表；未自动保存。Ctrl+Z 可按对象撤销。" +
                (skipped.Count == 0 ? "" : "\r\n需核对/保留的对象：\r\n" + string.Join("\r\n", skipped.Distinct()));
        }

        private static Word.Range FindObjectCaption(Word.Paragraph paragraph, bool figure)
        {
            for (int count = 0; count < 32; count++) {
                paragraph = figure ? paragraph.Next() : paragraph.Previous();
                if (paragraph == null) return null;
                Word.Range range = paragraph.Range;
                if (range.Tables.Count > 0 || range.InlineShapes.Count > 0) return null;
                string text = (range.Text ?? "").Trim();
                if (text.Length == 0) continue;
                string pattern = figure ? @"^(?:图\s*[A-Z0-9]|(?:Figure|Fig\.)\s*[A-Z0-9])" : @"^(?:表\s*[A-Z0-9]|Table\s*[A-Z0-9])";
                return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase) ? range : null;
            }
            return null;
        }

        private static float UsableObjectWidth(Word.Range range, bool picture)
        {
            Word.PageSetup page = range.Sections[1].PageSetup;
            if (page.TextColumns.Count != 1 || page.Gutter != 0) throw new InvalidDataException("分栏或装订线布局需人工核对。");
            float width = page.PageWidth - page.LeftMargin - page.RightMargin;
            if (picture) width -= range.ParagraphFormat.LeftIndent + range.ParagraphFormat.RightIndent + range.ParagraphFormat.FirstLineIndent;
            if (width <= 0 || width > 1584) throw new InvalidDataException("无法确认当前版心宽度。");
            return width;
        }

        private static void FormatPicture(Word.InlineShape picture, ObjectFormattingSettings options, Action mark)
        {
            float oldWidth = picture.Width, oldHeight = picture.Height;
            if (options.FitPictures) {
                float limit = UsableObjectWidth(picture.Range, true);
                if (oldWidth > limit + 0.1) {
                    var locked = picture.LockAspectRatio;
                    try {
                        picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse; mark();
                        picture.Width = limit; picture.Height = oldHeight * limit / oldWidth;
                    } finally { picture.LockAspectRatio = locked; }
                    if (Math.Abs(picture.Width - limit) > 0.1 || Math.Abs(picture.Height - oldHeight * limit / oldWidth) > 0.1)
                        throw new InvalidDataException("图片宽高回读与目标不一致。");
                }
            }
            if (options.CenterPictures) {
                picture.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; mark();
                if (picture.Range.ParagraphFormat.Alignment != Word.WdParagraphAlignment.wdAlignParagraphCenter)
                    throw new InvalidDataException("图片居中未生效。");
            }
        }

        private static void FormatTable(Word.Table table, ObjectFormattingSettings options, Action mark, List<string> skipped)
        {
            if (options.FitTables) {
                if (table.AllowAutoFit) skipped.Add("自动调整列宽表格：表内格式已处理，未强制更改列宽模式");
                else {
                    float limit = UsableObjectWidth(table.Range, false);
                    float[] before = ReadTableWidths(table);
                    if (before.Sum(width => (double)width) > limit + 0.1) {
                        WriteTableWidths(table, WordTableGeometry.ScaleColumns(before, limit), mark);
                        table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints; mark();
                        table.PreferredWidth = limit;
                        if (Math.Abs(ReadTableWidths(table).Sum(width => (double)width) - limit) > 0.1)
                            throw new InvalidDataException("表格列宽回读与版心不一致。");
                    }
                }
            }
            if (options.CenterTables) { table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter; mark(); }
            if (options.TableFont) {
                Word.Font font = table.Range.Font;
                font.NameFarEast = "宋体"; mark(); font.NameAscii = "Times New Roman"; font.NameOther = "Times New Roman";
                font.Size = options.TableSize;
                if (font.NameFarEast != "宋体" || font.NameAscii != "Times New Roman" || font.NameOther != "Times New Roman" ||
                    Math.Abs(font.Size - options.TableSize) > 0.01) throw new InvalidDataException("表内字体/字号回读不一致。");
            }
            if (options.TableAlignment) {
                table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; mark();
                table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                if (table.Range.ParagraphFormat.Alignment != Word.WdParagraphAlignment.wdAlignParagraphCenter ||
                    table.Range.Cells.VerticalAlignment != Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
                    throw new InvalidDataException("单元格对齐回读不一致。");
            }
            if (options.ThreeLine) {
                table.Borders.Enable = 0; mark();
                SetLine(table.Borders[Word.WdBorderType.wdBorderTop], Word.WdLineWidth.wdLineWidth150pt);
                SetLine(table.Borders[Word.WdBorderType.wdBorderBottom], Word.WdLineWidth.wdLineWidth150pt);
                SetLine(table.Rows[1].Borders[Word.WdBorderType.wdBorderBottom], Word.WdLineWidth.wdLineWidth075pt);
                if (table.Borders[Word.WdBorderType.wdBorderLeft].LineStyle != Word.WdLineStyle.wdLineStyleNone ||
                    table.Borders[Word.WdBorderType.wdBorderRight].LineStyle != Word.WdLineStyle.wdLineStyleNone ||
                    table.Rows[1].Borders[Word.WdBorderType.wdBorderBottom].LineWidth != Word.WdLineWidth.wdLineWidth075pt)
                    throw new InvalidDataException("三线表边框回读不一致。");
            }
            if (options.RepeatHeader) { table.Rows[1].HeadingFormat = -1; mark(); }
        }

        private static void SetLine(Word.Border border, Word.WdLineWidth width)
        { border.LineStyle = Word.WdLineStyle.wdLineStyleSingle; border.LineWidth = width; border.Color = Word.WdColor.wdColorBlack; }

        private static void FormatCaption(Word.Range caption, bool figure, Action mark)
        {
            caption.Font.NameFarEast = "宋体"; mark(); caption.Font.Size = 10.5F;
            Word.ParagraphFormat format = caption.ParagraphFormat;
            format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 6; format.SpaceAfter = figure ? 12 : 6;
            if (caption.Font.NameFarEast != "宋体" || Math.Abs(caption.Font.Size - 10.5F) > 0.01 ||
                format.Alignment != Word.WdParagraphAlignment.wdAlignParagraphCenter ||
                format.LineSpacingRule != Word.WdLineSpacing.wdLineSpaceSingle ||
                Math.Abs(format.SpaceBefore - 6) > 0.01 || Math.Abs(format.SpaceAfter - (figure ? 12 : 6)) > 0.01)
                throw new InvalidDataException("题注格式回读不一致。");
        }
    }
}
