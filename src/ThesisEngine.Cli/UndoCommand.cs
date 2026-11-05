using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>
/// 撤销最近一次格式修改。
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
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli undo <file.docx>");
            return 2;
        }

        var path = args[0];
        if (!File.Exists(path))
        {
            _error.WriteLine($"找不到文件：{path}");
            return 1;
        }

        try
        {
            _undo.Undo(path);
            _output.WriteLine($"✅ 已撤销最近一次修改，恢复到修改前的状态：{path}");
            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"撤销失败：{ex.Message}");
            return 1;
        }
    }
}
