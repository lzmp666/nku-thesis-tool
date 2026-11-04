using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public class ChatServiceTests
{
    [Fact]
    public async Task LocalRuleChatService_WithHeadingQuestion_ReturnsCorrectAnswer()
    {
        var service = new LocalRuleChatService();

        var answer = await service.AskAsync("一级标题用几号字");

        Assert.Contains("黑体三号", answer);
    }

    [Fact]
    public async Task LocalRuleChatService_WithBodyQuestion_ReturnsCorrectAnswer()
    {
        var service = new LocalRuleChatService();

        var answer = await service.AskAsync("正文用什么字体");

        Assert.Contains("宋体小四", answer);
    }

    [Fact]
    public async Task LocalRuleChatService_WithLineSpacingQuestion_ReturnsCorrectAnswer()
    {
        var service = new LocalRuleChatService();

        var answer = await service.AskAsync("行距多少");

        Assert.Contains("1.5 倍", answer);
    }

    [Fact]
    public async Task LocalRuleChatService_UnknownQuestion_ReturnsFallback()
    {
        var service = new LocalRuleChatService();

        var answer = await service.AskAsync("帮我润色这段话");

        Assert.Contains("格式相关", answer);
    }

    [Fact]
    public async Task LocalRuleChatService_EmptyQuestion_ReturnsPrompt()
    {
        var service = new LocalRuleChatService();

        var answer = await service.AskAsync("");

        Assert.Contains("请输入", answer);
    }
}
