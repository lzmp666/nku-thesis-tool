#nullable disable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;

namespace ThesisAddIn.VSTO.Engine
{
    // Keeps Word COM on its STA. Each write must close its own undo record and
    // restore ScreenUpdating before returning; no COM transaction spans an await.
    internal static class CooperativeRepairRunner
    {
        internal static async Task RunAsync<T>(IReadOnlyList<T> work, Action validate,
            Action<T> write, Func<Task> yield, Action<int> progress = null)
        {
            if (work == null || validate == null || write == null || yield == null)
                throw new ArgumentNullException();
            var slice = Stopwatch.StartNew();
            int count = 0;
            validate();
            for (int index = 0; index < work.Count; index++)
            {
                write(work[index]);
                count++;
                // Also yield after the LAST item: callers may immediately start
                // another rule group. Otherwise many one-location rules form one
                // long synchronous block even though every group awaits this method.
                if (index + 1 < work.Count && count < 4 && slice.ElapsedMilliseconds < 40) continue;
                progress?.Invoke(index + 1);
                await yield();
                // A user edit/cancellation during the yield must stop BEFORE
                // touching another range. The caller owns targeted rollback.
                validate();
                count = 0;
                slice.Restart();
            }
        }
    }
}
