#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Ui
{
    // One session for the searchable issue index and its Word-anchored annotations.
    internal sealed class IssueReviewState
    {
        internal readonly List<EngineDiagnostic> Rows = new List<EngineDiagnostic>();
        internal readonly HashSet<string> Repaired = new HashSet<string>(StringComparer.Ordinal);
        internal readonly HashSet<string> Checked = new HashSet<string>(StringComparer.Ordinal);
        internal bool BatchMode { get; private set; }
        internal bool BatchUndoMode { get; private set; }
        internal string SelectedKey { get; private set; } = string.Empty;
        internal string SelectedLocation { get; private set; } = string.Empty;
        internal event EventHandler? Changed;
        internal event EventHandler? SelectionChanged;
        internal static string Key(EngineDiagnostic? issue) => issue == null ? string.Empty :
            (issue.RuleId ?? "") + "\n" + (issue.Expected ?? "") + "\n" + (issue.FixLevel ?? "");
        internal bool IsRepaired(EngineDiagnostic issue) => Repaired.Contains(Key(issue));
        internal bool CanRepair(EngineDiagnostic issue) => issue != null && !IsRepaired(issue)
            && string.Equals(issue.FixLevel, "auto", StringComparison.OrdinalIgnoreCase);
        internal bool CanSelectForBatch(EngineDiagnostic issue) => BatchUndoMode ? IsRepaired(issue) : CanRepair(issue);
        internal void Refresh(IEnumerable<EngineDiagnostic> active)
        {
            var incoming = (active ?? Enumerable.Empty<EngineDiagnostic>()).Where(x => x != null)
                .GroupBy(Key).ToDictionary(group => group.Key, group => Merge(group), StringComparer.Ordinal);
            // An external Word undo or manual edit may make a repaired issue active again.
            foreach (string key in incoming.Keys) Repaired.Remove(key);
            var rows = new List<EngineDiagnostic>();
            foreach (EngineDiagnostic previous in Rows)
            {
                string key = Key(previous);
                if (incoming.TryGetValue(key, out EngineDiagnostic? next)) { rows.Add(next); incoming.Remove(key); }
                else if (Repaired.Contains(key)) rows.Add(previous);
            }
            rows.AddRange(incoming.Values);
            Rows.Clear(); Rows.AddRange(rows);
            Checked.RemoveWhere(key => !Rows.Any(row => Key(row) == key && CanSelectForBatch(row)));
            Changed?.Invoke(this, EventArgs.Empty);
        }
        private static EngineDiagnostic Merge(IEnumerable<EngineDiagnostic> group)
        {
            EngineDiagnostic first=group.First();
            var occurrences=group.SelectMany(x=>x.Occurrences!=null && x.Occurrences.Count>0
                ? x.Occurrences : new List<EngineDiagnosticOccurrence> { new EngineDiagnosticOccurrence { Location=x.Location, Actual=x.Actual } })
                .GroupBy(x=>x.Location).Select(x=>x.First()).ToList();
            return new EngineDiagnostic { RuleId=first.RuleId, Expected=first.Expected, FixLevel=first.FixLevel,
                Severity=first.Severity, Message=first.Message, Location=occurrences[0].Location,
                Actual=string.Join("；",occurrences.Select(x=>x.Actual).Distinct()), Occurrences=occurrences };
        }
        internal void Complete(IEnumerable<EngineDiagnostic> issues)
        {
            foreach (EngineDiagnostic issue in issues) { Repaired.Add(Key(issue)); Checked.Remove(Key(issue)); }
            Changed?.Invoke(this, EventArgs.Empty);
        }
        internal void Reopen(string key) => ReopenMany(new[]{key});
        internal void ReopenMany(IEnumerable<string> keys) {
            foreach (string key in keys) { Repaired.Remove(key); Checked.Remove(key); }
            Changed?.Invoke(this, EventArgs.Empty);
        }
        internal void SetBatch(bool value) => SetBatchMode(value, false);
        internal void SetBatchMode(bool value, bool undo) {
            if (BatchMode == value && BatchUndoMode == (value && undo)) return;
            BatchMode = value; BatchUndoMode = value && undo; Checked.Clear();
            Changed?.Invoke(this, EventArgs.Empty);
        }
        internal void Toggle(EngineDiagnostic issue)
        {
            if (!BatchMode || !CanSelectForBatch(issue)) return;
            string key = Key(issue); if (!Checked.Add(key)) Checked.Remove(key);
            Changed?.Invoke(this, EventArgs.Empty);
        }
        internal void SelectAll(bool value)
        {
            foreach (EngineDiagnostic row in Rows.Where(CanSelectForBatch))
                if (value) Checked.Add(Key(row)); else Checked.Remove(Key(row));
            Changed?.Invoke(this, EventArgs.Empty);
        }
        internal void Select(EngineDiagnostic issue, string location)
        {
            string key = Key(issue);
            if (key == SelectedKey && location == SelectedLocation) return;
            SelectedKey = key; SelectedLocation = location;
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
