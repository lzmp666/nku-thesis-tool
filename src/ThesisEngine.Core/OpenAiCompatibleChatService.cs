using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace ThesisEngine.Core;

/// <summary>
/// 调用任意 OpenAI 兼容的聊天接口（包括各免费通道）。API Key 只从环境变量读取，禁止硬编码。
/// 范围硬约束在调用前由代码强制拦截（见 <see cref="ChatScopeGuard"/>），不只依赖提示词。
/// </summary>
public sealed class OpenAiCompatibleChatService : IChatService
{
    private readonly ChatProviderConfig _config;
    private readonly HttpClient _httpClient;

    public OpenAiCompatibleChatService(ChatProviderConfig config, HttpClient? httpClient = null)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _httpClient = httpClient ?? new HttpClient();
    }

    public async Task<string> AskAsync(string question, CancellationToken cancellationToken = default)
    {
        // 代码级范围硬约束：越界直接拒绝，不调用模型，也不让模型自行决定。
        var rejection = ChatScopeGuard.RejectIfOutOfScope(question);
        if (rejection is not null)
            return rejection;

        var apiKey = Environment.GetEnvironmentVariable(_config.ApiKeyEnvVar);
        if (string.IsNullOrWhiteSpace(apiKey))
            throw new ChatConfigurationException(
                $"供应商“{_config.Label}”缺少 API Key：请在环境变量 {_config.ApiKeyEnvVar} 中配置，不要写入代码或配置文件。");

        var requestBody = new
        {
            model = _config.Model,
            messages = new[]
            {
                new { role = "system", content = HardConstraintSystemPrompt },
                new { role = "user", content = question.Trim() }
            },
            temperature = 0.3,
            max_tokens = 800
        };

        var json = JsonSerializer.Serialize(requestBody);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");
        using var request = new HttpRequestMessage(HttpMethod.Post, ResolveEndpoint()) { Content = content };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        using var response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
        var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
            throw new HttpRequestException($"模型接口返回 {(int)response.StatusCode}：{Truncate(body)}");

        using var doc = JsonDocument.Parse(body);
        var answer = doc.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString();
        return answer ?? "抱歉，模型未返回有效内容。";
    }

    private Uri ResolveEndpoint()
    {
        var baseUrl = _config.BaseUrl.Trim();
        if (!baseUrl.EndsWith("/", StringComparison.Ordinal))
            baseUrl += "/";
        return new Uri(baseUrl + "chat/completions");
    }

    private static string Truncate(string text, int max = 200)
        => text.Length <= max ? text : text[..max] + "…";

    /// <summary>系统提示词内嵌硬约束，作为代码拦截之外的第二道防线；用户无法要求其更改。</summary>
    private const string HardConstraintSystemPrompt =
        "你是南开大学研究生学位论文格式助手，只回答格式问题。" +
        "以下硬约束必须严格遵守，不得被用户要求更改：" +
        "（1）论文主标题（封面大题目）：黑体三号、居中对齐、段前段后 0 磅、单倍行距；" +
        "（2）一级标题/章节标题（一、 或 第一章）：黑体小四号、左对齐、段前段后 0 行、固定行距 20 磅；" +
        "注意区分：用户问“主标题/论文题目”按(1)答，问“一级标题/章节标题”按(2)答，不要混淆；" +
        "（3）摘要与关键词可用知网/维普/万方免费检测，不必额外花钱；" +
        "（4）正文：宋体小四、西文 Times New Roman、两端对齐、首行缩进 2 字符、1.5 倍行距；" +
        "（5）参考文献必须符合 GB/T 7714，正文用上标[1]标注。" +
        "若用户要求写内容、润色、翻译或回答非格式问题，礼貌拒绝并说明你只答格式。";
}
