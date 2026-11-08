using System.Text.Json;

namespace ThesisEngine.Core;

/// <summary>
/// 批量修改的撤销管理器。
/// 核心流程：先快照 → 再修改 → 撤销时回滚快照。
/// 每次备份都会在清单（manifest）中记录源文件、输出文件和将要做的改动，
/// 撤销时按清单弹出最近一次备份进行恢复，支持撤销最近一次批量修改。
/// </summary>
public sealed class UndoManager
{
    private const string ManifestFileName = "manifest.jsonl";

    private static readonly JsonSerializerOptions ManifestJsonOptions = new()
    {
        WriteIndented = false,
    };

    private readonly string _backupDir;
    private readonly string _manifestPath;

    public UndoManager(string? backupDir = null)
    {
        _backupDir = backupDir ?? Path.Combine(Path.GetTempPath(), "nku-thesis-tool-backups");
        Directory.CreateDirectory(_backupDir);
        _manifestPath = Path.Combine(_backupDir, ManifestFileName);
    }

    /// <summary>一次备份对应的清单记录。</summary>
    public sealed record UndoRecord(
        string BackupId,
        string BackupFileName,
        string SourcePath,
        string OutputPath,
        string ChangeDescription,
        string CreatedAt);

    /// <summary>
    /// 修改前先备份原文件，并记录本次批量修改的信息。
    /// </summary>
    /// <param name="filePath">要修改的源文件路径</param>
    /// <param name="outputPath">批量修改的输出文件路径（不传表示原地修改）</param>
    /// <param name="changeDescription">将要做的改动的描述，便于用户核对</param>
    /// <returns>备份 ID，撤销时用</returns>
    public string BackupBeforeModify(string filePath, string? outputPath = null, string? changeDescription = null)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("文件不存在", filePath);

        // 用时间戳生成唯一备份 ID（字典序即时间序）
        var backupId = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff");
        var fileName = Path.GetFileName(filePath);
        var backupPath = Path.Combine(_backupDir, $"{backupId}_{fileName}");

        File.Copy(filePath, backupPath, overwrite: true);

        AppendRecord(new UndoRecord(
            BackupId: backupId,
            BackupFileName: Path.GetFileName(backupPath),
            SourcePath: Path.GetFullPath(filePath),
            OutputPath: Path.GetFullPath(outputPath ?? filePath),
            ChangeDescription: changeDescription ?? "批量格式修改",
            CreatedAt: DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

        return backupId;
    }

    /// <summary>
    /// 撤销最近一次针对该文件的批量修改，恢复到修改前的状态。
    /// 撤销后对应的备份记录会被移除，再次撤销会回退到更早的一次备份。
    /// </summary>
    /// <param name="filePath">要恢复的文件路径（源文件或输出文件均可）</param>
    /// <returns>被撤销那条记录的备份 ID</returns>
    public string Undo(string filePath)
    {
        var target = Path.GetFullPath(filePath);
        var record = LoadRecords()
            .LastOrDefault(r =>
                string.Equals(r.OutputPath, target, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(r.SourcePath, target, StringComparison.OrdinalIgnoreCase));

        if (record == null)
            throw new InvalidOperationException("没有找到备份，无法撤销");

        var backupPath = Path.Combine(_backupDir, record.BackupFileName);
        if (!File.Exists(backupPath))
            throw new InvalidOperationException($"备份文件已丢失（{record.BackupFileName}），无法撤销");

        // 恢复的是“输出文件”的修改前状态（原地修改时即源文件本身）
        File.Copy(backupPath, record.OutputPath, overwrite: true);

        // 撤销后移除该条记录，避免重复恢复同一个备份
        RemoveRecord(record.BackupId);

        return record.BackupId;
    }

    /// <summary>按时间从新到旧列出所有备份记录。</summary>
    public IReadOnlyList<UndoRecord> ListBackups()
    {
        return LoadRecords();
    }

    /// <summary>清理旧备份，只保留最近 keepCount 次备份。</summary>
    public void CleanOldBackups(int keepCount = 10)
    {
        var records = LoadRecords();
        foreach (var oldRecord in records.Skip(keepCount))
        {
            var backupPath = Path.Combine(_backupDir, oldRecord.BackupFileName);
            if (File.Exists(backupPath))
                File.Delete(backupPath);
            RemoveRecord(oldRecord.BackupId);
        }
    }

    private void AppendRecord(UndoRecord record)
    {
        File.AppendAllText(_manifestPath, JsonSerializer.Serialize(record, ManifestJsonOptions) + Environment.NewLine);
    }

    private List<UndoRecord> LoadRecords()
    {
        var records = new List<UndoRecord>();
        if (!File.Exists(_manifestPath))
            return records;

        foreach (var line in File.ReadAllLines(_manifestPath))
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;
            try
            {
                var parsed = JsonSerializer.Deserialize<UndoRecord>(line, ManifestJsonOptions);
                if (parsed != null)
                    records.Add(parsed);
            }
            catch (JsonException)
            {
                // 忽略损坏的清单行，保持其余备份可用
            }
        }

        // 清单按追加顺序即时间顺序；保险起见按备份 ID 再排一次（ID 字典序即时间序）
        return records.OrderBy(r => r.BackupId, StringComparer.Ordinal).ToList();
    }

    private void RemoveRecord(string backupId)
    {
        if (!File.Exists(_manifestPath))
            return;

        var remaining = File.ReadAllLines(_manifestPath)
            .Where(line =>
            {
                if (string.IsNullOrWhiteSpace(line))
                    return false;
                try
                {
                    return JsonSerializer.Deserialize<UndoRecord>(line, ManifestJsonOptions)?.BackupId != backupId;
                }
                catch (JsonException)
                {
                    return true;
                }
            })
            .ToList();

        File.WriteAllLines(_manifestPath, remaining);
    }
}
