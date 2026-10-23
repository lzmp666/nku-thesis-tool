using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>
    /// Maps a native Word comment's stable diagnostic fields to one pane item.
    /// Ambiguous legacy comments deliberately do not select an arbitrary item.
    /// </summary>
    public static class CommentDiagnosticMatcher
    {
        /// <summary>Read the new one-field-per-line comments, with legacy inline fallback.</summary>
        public static string ReadCommentField(string text, string label)
        {
            text = text ?? string.Empty;
            string prefix = Regex.Escape(label) + @"[:：][ \t]*";
            Match line = Regex.Match(text, @"(?:^|[\r\n])" + prefix + @"(?<value>[^\r\n]*)");
            if (line.Success)
            {
                string value = line.Groups["value"].Value;
                int marker = value.IndexOf("[NKU-TOOL-ANNOTATION:v1]", StringComparison.Ordinal);
                if (marker >= 0) value = value.Substring(0, marker);
                return value.Trim().TrimEnd('。');
            }

            Match legacy = Regex.Match(text, prefix + @"(?<value>.*?)(?:[；;。\r\n]|$)");
            return legacy.Success ? legacy.Groups["value"].Value.Trim() : string.Empty;
        }

        public static int FindUniqueMatchIndex(
            IList<EngineDiagnostic> diagnostics,
            string ruleId,
            string expected,
            string actual,
            string location)
        {
            if (diagnostics == null || string.IsNullOrWhiteSpace(ruleId))
            {
                return -1;
            }

            int matchIndex = -1;
            for (int i = 0; i < diagnostics.Count; i++)
            {
                EngineDiagnostic diagnostic = diagnostics[i];
                if (diagnostic == null
                    || !string.Equals(diagnostic.RuleId, ruleId, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(diagnostic.Expected ?? string.Empty, expected ?? string.Empty, StringComparison.Ordinal)
                    || !string.Equals(diagnostic.Actual ?? string.Empty, actual ?? string.Empty, StringComparison.Ordinal))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(location) &&
                    !string.Equals(diagnostic.Location ?? string.Empty, location, StringComparison.Ordinal) &&
                    !ContainsOccurrence(diagnostic, location))
                    continue;

                if (matchIndex >= 0)
                {
                    return -1;
                }

                matchIndex = i;
            }

            return matchIndex;
        }

        private static bool ContainsOccurrence(EngineDiagnostic diagnostic, string location)
        {
            if (diagnostic.Occurrences == null) return false;
            foreach (EngineDiagnosticOccurrence occurrence in diagnostic.Occurrences)
            {
                if (occurrence != null && string.Equals(occurrence.Location, location, StringComparison.Ordinal))
                    return true;
            }
            return false;
        }
    }
}

