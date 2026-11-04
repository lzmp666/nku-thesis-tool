using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ThesisAddIn.VSTO.Engine
{
    internal static class WordTableGeometry
    {
        internal static float[] ScaleColumns(float[] widths, double targetPoints)
        {
            ReadColumns(new[] { widths });
            if (double.IsNaN(targetPoints) || double.IsInfinity(targetPoints) || targetPoints <= 0 || targetPoints > 1584)
                throw new InvalidDataException("目标表格宽度无效。");
            int targetTwips = (int)Math.Round(targetPoints * 20, MidpointRounding.AwayFromZero);
            double total = widths.Sum(value => (double)value);
            double[] exact = widths.Select(value => value * targetTwips / total).ToArray();
            int[] twips = exact.Select(value => (int)Math.Floor(value)).ToArray();
            int remainder = targetTwips - twips.Sum();
            foreach (int index in Enumerable.Range(0, widths.Length).OrderByDescending(index => exact[index] - twips[index]).Take(remainder))
                twips[index]++;
            if (twips.Any(value => value <= 0)) throw new InvalidDataException("目标列宽过小，已停止缩放。");
            return twips.Select(value => value / 20f).ToArray();
        }

        internal static bool EquivalentDimensions(double[] expected, double[] actual)
        {
            // Word lays out table widths asynchronously and stores dimensions in
            // twips. Do not confuse that sub-twip normalization with a user edit.
            return expected != null && actual != null && expected.Length == actual.Length &&
                expected.Zip(actual, (left, right) => !double.IsNaN(left) && !double.IsNaN(right) &&
                    !double.IsInfinity(left) && !double.IsInfinity(right) && Math.Abs(left - right) <= 0.1).All(equal => equal);
        }

        // Read concrete row/cell widths instead of Word.Table.Uniform, whose
        // aggregate flag can be false for ordinary compatibility-mode tables.
        internal static float[] ReadColumns(IReadOnlyList<float[]> rows)
        {
            if (rows == null || rows.Count == 0 || rows[0] == null || rows[0].Length == 0)
                throw new InvalidDataException("表格为空，不能安全缩放。");
            float[] widths = rows[0];
            foreach (float[] row in rows)
            {
                if (row == null || row.Length != widths.Length)
                    throw new InvalidDataException("表格单元格数不一致，已停止缩放。");
                for (int index = 0; index < row.Length; index++)
                    if (float.IsNaN(row[index]) || float.IsInfinity(row[index]) || row[index] <= 0 || row[index] > 1584 ||
                        Math.Abs(row[index] - widths[index]) > 0.1)
                        throw new InvalidDataException("表格各行列宽不一致或无法解析，已停止缩放。");
            }
            return widths.ToArray();
        }
    }
}
