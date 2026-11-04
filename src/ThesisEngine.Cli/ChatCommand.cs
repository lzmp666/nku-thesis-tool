using ThesisEngine.Core;

namespace ThesisEngine.Cli;

public sealed class ChatCommand
{
    private readonly IChatService _chatService;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public ChatCommand(
        IChatService? chatService = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _chatService = chatService ?? new LocalRuleChatService();
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        if (args.Length == 0)
        {
            _error.WriteLine("用法：ThesisEngine.Cli chat 你的问题");
            _error.WriteLine("示例：ThesisEngine.Cli chat 一级标题用几号字");
            return 2;
        }

        var question = string.Join(" ", args);
        _output.WriteLine($"你问：{question}");

        try
        {
            var answer = _chatService.AskAsync(question).GetAwaiter().GetResult();
            _output.WriteLine($"AI 回答：{answer}");
            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"AI 暂不可用：{ex.Message}");
            return 1;
        }
    }
}
