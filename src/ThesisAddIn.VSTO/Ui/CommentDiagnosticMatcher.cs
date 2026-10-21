using System;
using System.Collections.Generic;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    /// <summary>
    /// Maps a native Word comment's stable diagnostic fields to one pane item.
    /// Ambiguous legacy comments deliberately do not select an arbitrary item.
    /// </summary>
    public static class CommentDiagnosticMatcher
    {
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
                    || !string.Equals(diagnostic.Actual ?? string.Empty, actual ?? string.Empty, StringComparison.Ordinal)
                    || (!string.IsNullOrEmpty(location)
                        && !string.Equals(diagnostic.Location ?? string.Empty, location, StringComparison.Ordinal)))
                {
                    continue;
                }

                if (matchIndex >= 0)
                {
                    return -1;
                }

                matchIndex = i;
            }

            return matchIndex;
        }
    }
}

