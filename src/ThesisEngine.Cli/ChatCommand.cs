using ThesisEngine.Core;

namespace ThesisEngine.Cli;

public sealed class ChatCommand
{
    private readonly IChatService? _chatService;
    private readonly TextWriter _output;
    private readonly TextWriter _error;

    public ChatCommand(
        IChatService? chatService = null,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        _chatService = chatService;
        _output = output ?? Console.Out;
        _error = error ?? Console.Error;
    }

    public int Execute(string[] args)
    {
        var options = ParseArgs(args);

        if (options.Help)
        {
            PrintHelp();
            return 0;
        }

        if (string.IsNullOrWhiteSpace(options.Question))
        {
            PrintHelp();
            return 2;
        }

        IChatService service;
        try
        {
            service = _chatService ?? BuildService(options);
        }
        catch (ChatConfigurationException ex)
        {
            _error.WriteLine("配置错误：" + ex.Message);
            return 1;
        }

        _output.WriteLine($"你问：{options.Question}");

        try
        {
            var answer = service.AskAsync(options.Question).GetAwaiter().GetResult();
            _output.WriteLine($"AI 回答：{answer}");
            return 0;
        }
        catch (Exception ex)
        {
            _error.WriteLine($"AI 暂不可用：{ex.Message}");
            return 1;
        }
    }

    private static ChatOptions ParseArgs(string[] args)
    {
        var options = new ChatOptions();
        var questionTokens = new List<string>();
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg is "--help" or "-h")
            {
                options.Help = true;
            }
            else if (arg is "--local" or "-l")
            {
                options.Local = true;
            }
            else if (arg is "--config" or "-c")
            {
                if (i + 1 < args.Length)
                    options.ConfigPath = args[++i];
            }
            else
            {
                questionTokens.Add(arg);
            }
        }

        options.Question = string.Join(" ", questionTokens).Trim();
        return options;
    }

    private static IChatService BuildService(ChatOptions options)
    {
        if (options.Local)
            return new FormatConstraintChatService();

        var configuration = options.ConfigPath is not null
            ? ChatConfiguration.LoadFrom(options.ConfigPath)
            : ChatConfiguration.Load();

        return new ChatModelRouter(configuration);
    }

    private void PrintHelp()
    {
        _output.WriteLine("用法：ThesisEngine.Cli chat [选项] 你的问题");
        _output.WriteLine();
        _output.WriteLine("选项：");
        _output.WriteLine("  --local, -l          使用离线规则模式（无需任何 API Key，强制格式硬约束）");
        _output.WriteLine("  --config, -c <路径>   指定聊天模型配置文件（默认 chat.models.json）");
        _output.WriteLine("  --help, -h           显示本帮助");
        _output.WriteLine();
        _output.WriteLine("示例：");
        _output.WriteLine("  ThesisEngine.Cli chat 一级标题用几号字");
        _output.WriteLine("  ThesisEngine.Cli chat --local 正文用什么字体");
        _output.WriteLine("  ThesisEngine.Cli chat --config ./chat.models.json 参考文献怎么写");
        _output.WriteLine();
        _output.WriteLine("模型配置（免费通道，API Key 仅从环境变量读取，禁止硬编码）：");
        _output.WriteLine("  1) 复制 chat.models.json，按需启用供应商（SiliconFlow / DeepSeek / Kimi 等）。");
        _output.WriteLine("  2) 设置对应环境变量，例如：");
        _output.WriteLine("       set SILICONFLOW_API_KEY=sk-xxxx       (Windows)");
        _output.WriteLine("       export SILICONFLOW_API_KEY=sk-xxxx    (Linux/macOS)");
        _output.WriteLine("  3) 至少启用一个供应商并配置其 Key 后即可运行；全部失败时自动降级到其它免费通道。");
        _output.WriteLine("  详细说明见 docs/chat-models.md。");
    }

    private sealed class ChatOptions
    {
        public bool Help { get; set; }
        public bool Local { get; set; }
        public string? ConfigPath { get; set; }
        public string Question { get; set; } = string.Empty;
    }
}
