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
        internal static string TableRejection(string flatOpc)
        {
            var xml = XDocument.Parse(flatOpc);
            var tables = xml.Descendants(W + "tbl").ToArray();
            if (tables.Length != 1) return "嵌套表格";
            var table = tables[0];
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
