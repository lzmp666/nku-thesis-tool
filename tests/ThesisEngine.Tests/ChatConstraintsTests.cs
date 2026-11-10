using System.Collections.Generic;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// issue #113 验收：20 条测试问题（10 条格式类必须带对齐约束 + 10 条随意类必须明确拒绝），
/// 以及配置报错、模型降级等行为。
/// </summary>
public sealed class ChatConstraintsTests
{
    private static readonly string[] AlignmentKeywords = { "居中", "左对齐", "两端对齐" };

    private static bool HasAlignment(string answer)
        => AlignmentKeywords.Any(answer.Contains);

    // 10 条格式类问题：必须命中格式知识库，且回答包含对齐硬约束。
    [Theory]
    [InlineData("一级标题用几号字")]
    [InlineData("论文题目应该怎么排")]
    [InlineData("主标题字体字号")]
    [InlineData("一、这种章节标题格式")]
    [InlineData("二级标题用什么字体")]
    [InlineData("章标题怎么设置")]
    [InlineData("摘要部分怎么排版")]
    [InlineData("关键词放哪里")]
    [InlineData("正文用什么字体和字号")]
    [InlineData("图表怎么排版")]
    public async Task FormatQuestions_ReturnAnswerWithAlignmentConstraint(string question)
    {
        var service = new FormatConstraintChatService();
        var answer = await service.AskAsync(question);

        Assert.DoesNotContain("我只能回答", answer);
        Assert.True(HasAlignment(answer), $"回答未包含对齐硬约束：{answer}");
    }

    // 10 条随意类问题：必须被范围硬约束拒绝并给出明确理由。
    [Theory]
    [InlineData("帮我写一段论文内容")]
    [InlineData("把这句话翻译成英文")]
    [InlineData("讲个笑话听听")]
    [InlineData("帮我写个快排算法")]
    [InlineData("今天天气怎么样")]
    [InlineData("帮我润色这段文字")]
    [InlineData("推荐几本科幻小说")]
    [InlineData("帮我算一下 123 乘以 456")]
    [InlineData("简单解释一下相对论")]
    [InlineData("帮我写封求职邮件")]
    public async Task RandomQuestions_AreRejectedWithClearReason(string question)
    {
        var service = new FormatConstraintChatService();
        var answer = await service.AskAsync(question);

        Assert.Contains("我只能回答", answer);
    }

    [Fact]
    public async Task CitationQuestion_ReturnsGb7714Answer()
    {
        var service = new FormatConstraintChatService();
        var answer = await service.AskAsync("参考文献怎么写");

        Assert.Contains("GB/T 7714", answer);
        Assert.DoesNotContain("我只能回答", answer);
    }

    [Fact]
    public async Task InvalidCitation_IsRejectedWithClearError()
    {
        var service = new FormatConstraintChatService();
        var answer = await service.AskAsync("帮我检查这条引用 [张三 论文 北京]");

        Assert.Contains("不符合 GB/T 7714", answer);
    }

    [Fact]
    public void Router_WithoutUsableProvider_ThrowsConfigurationException()
    {
        Assert.Throws<ChatConfigurationException>(() => new ChatModelRouter(new ChatConfiguration()));
    }

    [Fact]
    public async Task Router_DegradesAcrossThreeLayers_WhenAllProvidersFail()
    {
        var envVars = new[] { "TEST_SF_KEY_113", "TEST_DS_KEY_113", "TEST_MK_KEY_113" };
        foreach (var envVar in envVars)
            Environment.SetEnvironmentVariable(envVar, "dummy");

        try
        {
            var config = new ChatConfiguration
            {
                Providers = new List<ChatProviderConfig>
                {
                    new() { Name = "a", Label = "A", BaseUrl = "http://127.0.0.1:9/v1", Model = "m", ApiKeyEnvVar = "TEST_SF_KEY_113", Enabled = true, Priority = 1 },
                    new() { Name = "b", Label = "B", BaseUrl = "http://127.0.0.1:9/v1", Model = "m", ApiKeyEnvVar = "TEST_DS_KEY_113", Enabled = true, Priority = 2 },
                    new() { Name = "c", Label = "C", BaseUrl = "http://127.0.0.1:9/v1", Model = "m", ApiKeyEnvVar = "TEST_MK_KEY_113", Enabled = true, Priority = 3 }
                }
            };

            var router = new ChatModelRouter(config);
            var answer = await router.AskAsync("一级标题用几号字");

            Assert.Contains("层级", answer);
        }
        finally
        {
            foreach (var envVar in envVars)
                Environment.SetEnvironmentVariable(envVar, null);
        }
    }
}
