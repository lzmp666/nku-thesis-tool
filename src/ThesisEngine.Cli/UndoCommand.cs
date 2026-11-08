using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 撤销最近一次批量格式修改。
/// 用法：ThesisEngine.Cli undo &lt;file.docx&gt;　或　ThesisEngine.Cli undo --list [&lt;file.docx&gt;]
/// </summary>
public sealed class UndoCommand
{
    private readonly UndoManager _undo;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public UndoCommand(
        UndoManager? undo = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _undo = undo ?? new UndoManager();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        var listOnly = args.Length > 0 && string.Equals(args[0], "--list", StringComparison.OrdinalIgnoreCase);
        var path = listOnly ? args.Skip(1).FirstOrDefault() : args.FirstOrDefault();

        if (listOnly)
            return ExecuteList(path);

        if (string.IsNullOrEmpty(path))
        {
            _error.WriteLine("用法：ThesisEngine.Cli undo <file.docx>（或 undo --list 查看可撤销的备份）");
            return 2;
        }

        if (!File.Exists(path))
        {
            _error.WriteLine($"找不到文件：{path}");
            return 1;
        }

        try
        {
            var backupId = _undo.Undo(path);
            _output.WriteLine($"✅ 已撤销最近一次批量修改，恢复到修改前的状态：{path}");
            _output.WriteLine($"   （回滚的备份 ID：{backupId}）");
            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"撤销失败：{ex.Message}");
            return 1;
        }
    }

    private int ExecuteList(string? path)
    {
        var records = _undo.ListBackups();
        if (!string.IsNullOrEmpty(path))
        {
            var target = Path.GetFullPath(path);
            records = records
                .Where(r => string.Equals(r.OutputPath, target, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(r.SourcePath, target, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        if (records.Count == 0)
        {
            _output.WriteLine("没有可撤销的备份。批量修改时会自动创建快照。");
            return 0;
        }

        _output.WriteLine("可撤销的批量修改（从最近到最早）：");
        foreach (var record in records)
        {
            _output.WriteLine($"  [{record.CreatedAt}] {record.BackupId}");
            _output.WriteLine($"    改动：{record.ChangeDescription}");
            _output.WriteLine($"    源文件：{record.SourcePath}");
            if (!string.Equals(record.OutputPath, record.SourcePath, StringComparison.OrdinalIgnoreCase))
                _output.WriteLine($"    输出文件：{record.OutputPath}");
        }

        return 0;
    }
}
