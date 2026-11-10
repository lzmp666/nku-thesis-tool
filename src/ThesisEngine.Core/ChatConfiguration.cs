using System.Linq;
using System.Text.Json;

namespace ThesisEngine.Core;

/// <summary>
/// 单个模型供应商的配置。API Key 只允许通过环境变量名引用，禁止在配置文件或代码中硬编码密钥本身。
/// </summary>
public sealed class ChatProviderConfig
{
    /// <summary>供应商内部名称（如 siliconflow）。</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>展示名称，仅用于日志与提示。</summary>
    public string Label { get; set; } = string.Empty;

    /// <summary>OpenAI 兼容的 API 基地址，调用时自动拼接 chat/completions。</summary>
    public string BaseUrl { get; set; } = string.Empty;

    /// <summary>模型名称，由供应商决定。</summary>
    public string Model { get; set; } = string.Empty;

    /// <summary>读取 API Key 的环境变量名（如 SILICONFLOW_API_KEY）。不要把密钥值写进这里。</summary>
    public string ApiKeyEnvVar { get; set; } = string.Empty;

    /// <summary>是否启用该供应商。</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>优先级，数值越小越先用。</summary>
    public int Priority { get; set; } = 100;
}

/// <summary>
/// 聊天模型配置：从默认位置（环境变量 THESIS_CHAT_CONFIG → 程序目录 → 当前目录）加载 chat.models.json。
/// API Key 一律从环境变量读取，配置文件中只保留环境变量名。
/// </summary>
public sealed class ChatConfiguration
{
    /// <summary>已启用的供应商列表。</summary>
    public IReadOnlyList<ChatProviderConfig> Providers { get; init; } = Array.Empty<ChatProviderConfig>();

    /// <summary>是否在没有任何模型 Key 时允许回退到本地离线规则（默认关闭）。</summary>
    public bool AllowLocalFallback { get; init; }

    /// <summary>空配置（无任何供应商）。</summary>
    public static ChatConfiguration Empty { get; } = new();

    /// <summary>按默认顺序解析配置路径并加载；找不到则返回 <see cref="Empty"/>。</summary>
    public static ChatConfiguration Load()
    {
        var path = ResolveConfigPath();
        return path is null ? Empty : LoadFrom(path);
    }

    /// <summary>从指定路径加载配置文件。</summary>
    public static ChatConfiguration LoadFrom(string path)
    {
        if (!File.Exists(path))
            throw new ChatConfigurationException($"找不到聊天模型配置文件：{path}");
        return Parse(File.ReadAllText(path));
    }

    /// <summary>解析配置 JSON。</summary>
    public static ChatConfiguration Parse(string json)
    {
        var file = JsonSerializer.Deserialize<ConfigFile>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        var providers = (file?.Providers ?? new List<ChatProviderConfig>())
            .ConvertAll(p => p ?? new ChatProviderConfig());
        return new ChatConfiguration
        {
            Providers = providers,
            AllowLocalFallback = file?.AllowLocalFallback ?? false
        };
    }

    /// <summary>返回已启用、基地址与 Key 环境变量齐全、且环境变量中确实取到 Key 的供应商，按优先级排序。</summary>
    public IReadOnlyList<ChatProviderConfig> ResolveUsableProviders()
    {
        return Providers
            .Where(p => p.Enabled
                        && !string.IsNullOrWhiteSpace(p.BaseUrl)
                        && !string.IsNullOrWhiteSpace(p.ApiKeyEnvVar)
                        && !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(p.ApiKeyEnvVar)))
            .OrderBy(p => p.Priority)
            .ThenBy(p => p.Name, StringComparer.Ordinal)
            .ToArray();
    }

    private static string? ResolveConfigPath()
    {
        var fromEnv = Environment.GetEnvironmentVariable("THESIS_CHAT_CONFIG");
        if (!string.IsNullOrWhiteSpace(fromEnv) && File.Exists(fromEnv))
            return fromEnv;

        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "chat.models.json"),
            Path.Combine(Environment.CurrentDirectory, "chat.models.json")
        };
        foreach (var candidate in candidates)
            if (File.Exists(candidate))
                return candidate;

        return null;
    }

    private sealed class ConfigFile
    {
        public List<ChatProviderConfig>? Providers { get; set; }
        public bool AllowLocalFallback { get; set; }
    }
}

/// <summary>配置缺失或无效时抛出，由 ChatCommand 捕获并给出清晰提示（不崩溃）。</summary>
public sealed class ChatConfigurationException : Exception
{
    public ChatConfigurationException(string message) : base(message)
    {
    }
}
