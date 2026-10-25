#nullable disable
using System;
using System.Linq;
using System.Xml.Linq;

namespace ThesisAddIn.VSTO.Engine
{
    internal sealed class ObjectFormattingSettings
    {
        internal bool CenterPictures = true, FitPictures = true, CenterTables = true,
            FitTables = true, TableFont = true, TableAlignment = true, ThreeLine = true,
            RepeatHeader = true, Captions = true;
        internal float TableSize = 10.5F;
    }

    internal static class ObjectFormattingPolicy
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        internal static string TableRejection(string flatOpc, string targetText = null)
        {
            var xml = XDocument.Parse(flatOpc);
            XNamespace package = "http://schemas.microsoft.com/office/2006/xmlPackage";
            var main = xml.Descendants(package + "part").FirstOrDefault(part =>
                (string)part.Attribute(package + "name") == "/word/document.xml");
            // A Flat OPC selection can carry glossary/building-block dependencies
            // containing unrelated tables. Only inspect its document body.
            var scope = main?.Element(package + "xmlData") ?? xml.Root;
            var tables = scope.DescendantsAndSelf(W + "tbl").ToArray();
            if (targetText != null) {
                string text = targetText.Replace("\r", "").Replace("\a", "").Replace("\v", "");
                tables = tables.Where(table => !table.Ancestors(W + "tbl").Any() &&
                    string.Concat(table.Descendants(W + "t").Select(value => value.Value)) == text).ToArray();
            }
            if (tables.Length != 1) return "无法唯一匹配表格正文（匹配 " + tables.Length + " 张表）";
            var table = tables[0];
            if (table.Descendants(W + "tbl").Any()) return "嵌套表格";
            if (table.Descendants().Any(e => e.Name.Namespace == W &&
                new[] { "gridSpan", "vMerge", "hMerge", "gridBefore", "gridAfter", "ins", "del",
                    "tblPrChange", "tcPrChange", "fldChar", "fldSimple", "drawing", "pict", "sdt" }.Contains(e.Name.LocalName)))
                return "合并单元格、域、修订或嵌入对象";
            if (table.Descendants().Any(e => e.Name.NamespaceName.Contains("math"))) return "含公式的表格";
            if (table.Element(W + "tblPr")?.Element(W + "tblpPr") != null) return "浮动表格";
            var rows = table.Elements(W + "tr").ToArray();
            int columns = table.Element(W + "tblGrid")?.Elements(W + "gridCol").Count() ?? 0;
            if (rows.Length < 2 || columns == 0 || rows.Sum(row => row.Elements(W + "tc").Count()) > 2000 ||
                rows.Any(row => row.Elements(W + "tc").Count() != columns)) return "表头或矩形结构不明确，或表格过大";
            return null;
        }
    }
}
