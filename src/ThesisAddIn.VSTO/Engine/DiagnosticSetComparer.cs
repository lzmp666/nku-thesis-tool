using System;
using System.Collections.Generic;
using System.Linq;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Engine
{
    internal static class DiagnosticSetComparer
    {
        internal static bool Equivalent(IEnumerable<EngineDiagnostic> expected, IEnumerable<EngineDiagnostic> actual)
        {
            return expected.Select(Signature).OrderBy(value => value, StringComparer.Ordinal).SequenceEqual(
                actual.Select(Signature).OrderBy(value => value, StringComparer.Ordinal));
        }
        private static string Signature(EngineDiagnostic row)
        {
            // Length-delimit fields: arbitrary user text cannot collide with separators.
            Func<string,string> field = value => (value ?? "").Length + ":" + (value ?? "");
            var occurrences = row.Occurrences != null && row.Occurrences.Count > 0 ? row.Occurrences :
                new List<EngineDiagnosticOccurrence> { new EngineDiagnosticOccurrence { Location = row.Location, Actual = row.Actual } };
            return field(row.RuleId) + field(row.Expected) + field(row.Actual) + field(row.FixLevel) + field(row.Severity) + field(row.Message) +
                string.Concat(occurrences.OrderBy(value => value.Location, StringComparer.Ordinal).Select(value => field(value.Location) + field(value.Actual)));
        }
    }
}
