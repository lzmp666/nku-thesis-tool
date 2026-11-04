using ThesisEngine.Cli;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public class ChatCommandTests
{
    [Fact]
    public void ChatCommand_WithNoArgs_ShowsUsageAndReturnsError()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var command = new ChatCommand(output: output, error: error);

        var result = command.Execute(Array.Empty<string>());

        Assert.Equal(2, result);
        Assert.Contains("用法", error.ToString());
    }

    [Fact]
    public void ChatCommand_WithHeadingQuestion_ReturnsCorrectAnswer()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var command = new ChatCommand(output: output, error: error);

        var result = command.Execute(new[] { "一级标题", "用几号字" });

        Assert.Equal(0, result);
        var outputText = output.ToString();
        Assert.Contains("你问：一级标题 用几号字", outputText);
        Assert.Contains("黑体三号", outputText);
    }

    [Fact]
    public void ChatCommand_WithBodyQuestion_ReturnsCorrectAnswer()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var command = new ChatCommand(output: output, error: error);

        var result = command.Execute(new[] { "正文", "用什么字体" });

        Assert.Equal(0, result);
        var outputText = output.ToString();
        Assert.Contains("宋体小四", outputText);
    }

    [Fact]
    public async Task LocalRuleChatService_UnknownQuestion_ReturnsFallback()
    {
        var service = new LocalRuleChatService();

        var answer = await service.AskAsync("帮我润色这段话");

        Assert.Contains("格式相关", answer);
    }
}
