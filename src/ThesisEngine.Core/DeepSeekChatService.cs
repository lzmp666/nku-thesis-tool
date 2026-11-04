using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace ThesisEngine.Core;

/// <summary>
/// DeepSeek 大模型聊天服务。
/// 调用 DeepSeek API 回答论文格式相关问题。
/// </summary>
public sealed class DeepSeekChatService : IChatService
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;
    private readonly string _model;

    public DeepSeekChatService(string apiKey, HttpClient? httpClient = null, string model = "deepseek-chat")
    {
        _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
        _model = model;
        _httpClient = httpClient ?? new HttpClient();
        _httpClient.BaseAddress = new Uri("https://api.deepseek.com/v1/");
    }

    public async Task<string> AskAsync(string question, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(question))
        {
            return "请输入你想问的格式问题。";
        }

        var systemPrompt = @"你是南开大学研究生论文格式助手。你的职责：
1. 只回答与论文格式相关的问题（字体、字号、行距、标题、图表、参考文献等）
2. 不回答论文内容、写作思路、文献综述等问题
3. 回答要简洁、准确，基于南开大学研究生学位论文格式要求
4. 如果问题与格式无关，礼貌地说明你只能回答格式问题";

        var requestBody = new
        {
            model = _model,
            messages = new[]
            {
                new { role = "system", content = systemPrompt },
                new { role = "user", content = question.Trim() }
            },
            temperature = 0.3,
            max_tokens = 500
        };

        var json = JsonSerializer.Serialize(requestBody);
        var content = new StringContent(json, Encoding.UTF8, "application/json");

        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);

        var response = await _httpClient.PostAsync("chat/completions", content, cancellationToken);
        var responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            return $"AI 服务暂时不可用（{response.StatusCode}），请稍后再试。";
        }

        using var doc = JsonDocument.Parse(responseBody);
        var answer = doc.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString();

        return answer ?? "抱歉，我暂时无法回答这个问题。";
    }
}
