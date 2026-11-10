using System.Linq;
using System.Text.Json;

namespace ThesisEngine.Core;

/// <summary>
/// 多模型供应商路由器：按优先级依次尝试已配置的免费通道。
/// 主通道失败后降级到另外 2 个免费通道，并逐层上报错误（共 3 层）。
/// 没有任何可用供应商时抛出 <see cref="ChatConfigurationException"/>（由上层捕获，不崩溃）。
/// </summary>
public sealed class ChatModelRouter : IChatService
{
    private readonly IReadOnlyList<ChatProviderConfig> _providers;
    private readonly HttpClient? _httpClient;

    public ChatModelRouter(ChatConfiguration configuration, HttpClient? httpClient = null)
    {
        if (configuration is null)
            throw new ArgumentNullException(nameof(configuration));

        _httpClient = httpClient;
        var usable = configuration.ResolveUsableProviders();
        if (usable.Count == 0)
            throw new ChatConfigurationException(BuildNoProviderMessage(configuration));

        _providers = usable;
    }

    public async Task<string> AskAsync(string question, CancellationToken cancellationToken = default)
    {
        // 先做代码级范围硬约束（所有通道共享，保证不被单个模型绕开）。
        var rejection = ChatScopeGuard.RejectIfOutOfScope(question);
        if (rejection is not null)
            return rejection;

        var layers = new List<string>();
        // 最多尝试 3 个通道（主 + 2 个降级），即“报 3 层”。
        foreach (var provider in _providers.Take(3))
        {
            try
            {
                var service = new OpenAiCompatibleChatService(provider, _httpClient);
                return await service.AskAsync(question, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is HttpRequestException or IOException or TaskCanceledException
                or OperationCanceledException or JsonException or ChatConfigurationException)
            {
                layers.Add($"[{provider.Label}] 失败：{ex.Message}");
            }
        }

        // 全部失败：汇总 3 层错误，明确告知，不崩溃。
        return "所有模型通道均不可用，已按以下层级依次失败：\n" + string.Join("\n", layers) +
               "\n请检查网络、对应环境变量中的 API Key，或改用 --local 离线模式。";
    }

    private static string BuildNoProviderMessage(ChatConfiguration configuration)
    {
        var enabled = configuration.Providers.Where(p => p.Enabled).ToArray();
        if (enabled.Length == 0)
            return "配置文件中没有任何已启用的模型供应商。请在 chat.models.json 中启用至少一个供应商，并配置对应的 API Key 环境变量，或使用 --local 离线模式。";

        var parts = enabled.Select(p => $"{p.Label}（需设置环境变量 {p.ApiKeyEnvVar}）").ToArray();
        return "未检测到可用的模型 API Key。已启用的供应商需要设置以下环境变量之一：" +
               string.Join("、", parts) +
               "。请配置后在环境变量中填入 Key（不要写进代码或配置文件），或使用 --local 离线模式。";
    }
}
