namespace ThesisEngine.Core;

/// <summary>
/// 批量修改的撤销管理器。
/// 修改前自动备份，撤销时恢复到修改前的状态。
/// </summary>
public sealed class UndoManager
{
    private readonly string _backupDir;

    public UndoManager(string? backupDir = null)
    {
        _backupDir = backupDir ?? Path.Combine(Path.GetTempPath(), "nku-thesis-tool-backups");
        Directory.CreateDirectory(_backupDir);
    }

    /// <summary>
    /// 修改前先备份原文件。
    /// </summary>
    /// <param name="filePath">要修改的文件路径</param>
    /// <returns>备份 ID，撤销时用</returns>
    public string BackupBeforeModify(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("文件不存在", filePath);

        // 用时间戳生成唯一备份 ID
        var backupId = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff");
        var fileName = Path.GetFileName(filePath);
        var backupPath = Path.Combine(_backupDir, $"{backupId}_{fileName}");

        File.Copy(filePath, backupPath, overwrite: true);

        return backupId;
    }

    /// <summary>
    /// 撤销最近一次修改，恢复到备份的状态。
    /// </summary>
    /// <param name="filePath">要恢复的文件路径</param>
    public void Undo(string filePath)
    {
        var fileName = Path.GetFileName(filePath);

        // 找最新的备份
        var latestBackup = Directory.GetFiles(_backupDir, $"*_{fileName}")
            .OrderByDescending(f => f)
            .FirstOrDefault();

        if (latestBackup == null)
        {
            throw new InvalidOperationException("没有找到备份，无法撤销");
        }

        File.Copy(latestBackup, filePath, overwrite: true);
    }

    /// <summary>
    /// 列出所有备份。
    /// </summary>
    public List<string> ListBackups()
    {
        return Directory.GetFiles(_backupDir)
            .OrderByDescending(f => f)
            .Select(Path.GetFileName)
            .ToList()!;
    }

    /// <summary>
    /// 清理所有旧备份。
    /// </summary>
    public void CleanOldBackups(int keepCount = 10)
    {
        var backups = Directory.GetFiles(_backupDir)
            .OrderByDescending(f => f)
            .ToList();

        foreach (var oldBackup in backups.Skip(keepCount))
        {
            File.Delete(oldBackup);
        }
    }
}
