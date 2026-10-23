using System;
using System.Globalization;
using System.IO;

namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>Map Word PIA values to the same values used by the OOXML reader.</summary>
    internal static class WordParagraphValueMapper
    {
        internal static string Alignment(int alignment)
        {
            switch (alignment)
            {
                case 0: return "left";
                case 1: return "center";
                case 2: return "right";
                case 3: return "both";
                case 4: return "distribute";
                default: throw new InvalidDataException("题注对齐为混合值或暂不支持的值；未执行修复。");
            }
        }

        internal static string LineSpacing(int rule, double points)
        {
            switch (rule)
            {
                case 0: return "single";
                case 1: return "multiple:360";
                case 2: return "multiple:480";
                case 3: case 4: case 5:
                    if (double.IsNaN(points) || double.IsInfinity(points) || points <= 0 || points > 1584)
                        throw new InvalidDataException("题注行距为混合值或无效值；未执行修复。");
                    // Word measures multiple spacing against a 12pt line: 18pt is
                    // 1.5 lines. OOXML auto spacing uses 240ths of a line; exact
                    // and atLeast use twips. Both conversions therefore use 20.
                    int value = (int)Math.Round(points * 20, MidpointRounding.AwayFromZero);
                    if (rule == 5 && value == 240) return "single";
                    string prefix = rule == 3 ? "atLeast:" : rule == 4 ? "exact:" : "multiple:";
                    return prefix + value.ToString(CultureInfo.InvariantCulture);
                default: throw new InvalidDataException("题注行距类型为混合值或暂不支持的值；未执行修复。");
            }
        }
    }
}
