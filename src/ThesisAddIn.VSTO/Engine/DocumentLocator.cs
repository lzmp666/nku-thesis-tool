using System;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>引擎给出的稳定位置标识，例如 word/document.xml#paragraph[2]/run[0]。</summary>
    public sealed class DocumentLocation
    {
        public string Part { get; set; }

        /// <summary>正文段落序号，从 0 开始；-1 表示该位置未给出段落。</summary>
        public int ParagraphIndex { get; set; }

        /// <summary>段落内文本片段序号，从 0 开始；-1 表示未给出。</summary>
        public int RunIndex { get; set; }
    }

    public static class DocumentLocator
    {
        private static readonly Regex LocationPattern = new Regex(
            @"^(?<part>[^#]+)#paragraph\[(?<paragraph>\d+)\](?:/run\[(?<run>\d+)\])?$",
            RegexOptions.Compiled);

        public static bool TryParse(string location, out DocumentLocation parsed, out string error)
        {
            parsed = null;
            error = null;

            if (string.IsNullOrWhiteSpace(location))
            {
                error = "这条诊断没有位置信息，无法定位。";
                return false;
            }

            Match match = LocationPattern.Match(location.Trim());
            if (!match.Success)
            {
                error = "无法识别的位置格式：" + location + "。当前支持“部件#paragraph[序号]/run[序号]”。";
                return false;
            }

            parsed = new DocumentLocation
            {
                Part = match.Groups["part"].Value,
                ParagraphIndex = int.Parse(match.Groups["paragraph"].Value),
                RunIndex = match.Groups["run"].Success ? int.Parse(match.Groups["run"].Value) : -1
            };
            return true;
        }

        /// <summary>
        /// 把段落序号映射到 Word 段落范围。注意：引擎只统计正文里的段落，
        /// 而 Word 的 Paragraphs 集合还包含表格单元格内的段落，因此含表格的
        /// 文档可能出现错位；此时本方法会给出提示而不是静默选中错误内容。
        /// </summary>
        public static bool TryGetParagraphRange(Word.Document document, DocumentLocation location, out Word.Range range, out string error)
        {
            range = null;
            error = null;

            if (document == null)
            {
                error = "当前没有打开的文档。";
                return false;
            }

            if (location == null || location.ParagraphIndex < 0)
            {
                error = "这条诊断没有可用的段落序号。";
                return false;
            }

            if (!string.Equals(location.Part, "word/document.xml", StringComparison.OrdinalIgnoreCase))
            {
                error = "当前仅支持正文中的段落定位，无法安全定位到“" + location.Part + "”。";
                return false;
            }

            // 引擎按 document.xml 的正文段落计数；Word 的 Paragraphs 还包括表格单元格段落。
            // 只靠序号无法在有表格的文档中保证对应关系，因此宁可停止定位，也不选错原文。
            if (document.Tables.Count > 0)
            {
                error = "当前文档含有表格，段落序号可能无法准确对应；为避免定位到错误内容，已停止本次定位。";
                return false;
            }

            int wordIndex = location.ParagraphIndex + 1; // Word 集合从 1 开始
            int paragraphCount = document.Paragraphs.Count;
            if (wordIndex < 1 || wordIndex > paragraphCount)
            {
                error = "诊断指向第 " + wordIndex + " 段，但当前文档只有 " + paragraphCount + " 段，已停止定位。";
                return false;
            }

            range = document.Paragraphs[wordIndex].Range;
            return true;
        }

        public static string Describe(DocumentLocation location)
        {
            if (location == null)
            {
                return "（无位置）";
            }

            string text = "第 " + (location.ParagraphIndex + 1) + " 段";
            if (location.RunIndex >= 0)
            {
                text += " 第 " + (location.RunIndex + 1) + " 个文本片段";
            }

            return text;
        }
    }
}
