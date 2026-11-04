namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>
    /// An install/update writes the current-user setting. A launcher that was already
    /// running can still pass the old process environment to a newly launched Word.
    /// Never let that stale value override an explicit current-user installation.
    /// </summary>
    internal static class EngineHostPathResolver
    {
        internal static string Select(string userHostPath, string processHostPath)
        {
            return string.IsNullOrWhiteSpace(userHostPath) ? processHostPath : userHostPath;
        }
    }
}
