# 聊天模型配置（issue #113）

`chat` 子命令用于回答南开大学研究生学位论文的**格式**问题。本文件说明如何为它接入模型、如何排查问题，以及每个配置参数的含义。

> 设计原则：**API Key 只允许从环境变量读取，禁止硬编码到代码或配置文件**。格式范围由代码强制约束，不只依赖提示词。

## 1. 三条硬约束（代码级，不可绕过）

1. **范围硬约束**：聊天只回答格式问题（字体、字号、对齐、行距、标题、图表、摘要、参考文献等）。任何写内容、润色、翻译或无关话题的请求都会被代码拦截并返回明确拒绝理由。
2. **格式硬约束**：标题黑体三号居中、章节标题黑体小四固定 20 磅行距、摘要关键词可用知网/维普/万方免费检测等，均写死在 `FormatKnowledge` 中，模型也无法更改。
3. **引用/标注约束**：参考文献必须符合 GB/T 7714，不合规会明确报错，不允许含糊带过。

## 2. 两种运行模式

| 模式 | 命令 | 是否需要 API Key | 说明 |
| --- | --- | --- | --- |
| 离线规则 | `chat --local 你的问题` | 否 | 走内置规则库，强制上述硬约束，适合无网络或快速验证 |
| 模型通道 | `chat 你的问题` | 是 | 走 `chat.models.json` 配置的免费模型通道，主通道失败自动降级 |

## 3. 配置文件 `chat.models.json`

默认读取顺序：环境变量 `THESIS_CHAT_CONFIG` 指定的路径 → 程序所在目录的 `chat.models.json` → 当前工作目录的 `chat.models.json`。

> 仓库内置一份默认 `chat.models.json`（位于 `src/ThesisEngine.Cli/`，构建时复制到输出目录），已列出 3 个免费通道 + 1 个 OpenAI 兼容接口模板。

### 顶层字段

| 字段 | 类型 | 含义 |
| --- | --- | --- |
| `allowLocalFallback` | bool | 没有任何模型 Key 时是否允许回退到本地离线规则。默认 `false`，即无 Key 时明确报错（不崩溃） |
| `providers` | 数组 | 模型供应商列表 |

### 每个 provider 的字段

| 字段 | 类型 | 含义 |
| --- | --- | --- |
| `name` | string | 内部标识，唯一即可 |
| `label` | string | 展示名称，仅用于日志与提示 |
| `baseUrl` | string | OpenAI 兼容的 API 基地址，调用时自动拼接 `chat/completions` |
| `model` | string | 模型名，由供应商决定 |
| `apiKeyEnvVar` | string | **读取 API Key 的环境变量名**。不要把密钥值写在这里 |
| `enabled` | bool | 是否启用该供应商 |
| `priority` | int | 优先级，数值越小越先用 |

### 默认配置示例

```json
{
  "allowLocalFallback": false,
  "providers": [
    {
      "name": "siliconflow",
      "label": "硅基流动 SiliconFlow（免费额度）",
      "baseUrl": "https://api.siliconflow.cn/v1",
      "model": "deepseek-ai/DeepSeek-V3",
      "apiKeyEnvVar": "SILICONFLOW_API_KEY",
      "enabled": true,
      "priority": 1
    },
    {
      "name": "deepseek",
      "label": "DeepSeek（免费额度）",
      "baseUrl": "https://api.deepseek.com/v1",
      "model": "deepseek-chat",
      "apiKeyEnvVar": "DEEPSEEK_API_KEY",
      "enabled": true,
      "priority": 2
    },
    {
      "name": "moonshot",
      "label": "Kimi / Moonshot（免费额度）",
      "baseUrl": "https://api.moonshot.cn/v1",
      "model": "moonshot-v1-8k",
      "apiKeyEnvVar": "MOONSHOT_API_KEY",
      "enabled": true,
      "priority": 3
    },
    {
      "name": "openai-compatible",
      "label": "OpenAI 兼容接口（自定义 base_url + model）",
      "baseUrl": "https://api.openai.com/v1",
      "model": "gpt-4o-mini",
      "apiKeyEnvVar": "OPENAI_API_KEY",
      "enabled": false,
      "priority": 4
    }
  ]
}
```

## 4. 怎么配（三步）

1. 复制并编辑 `chat.models.json`，按需启用供应商、调整 `baseUrl`/`model`/`apiKeyEnvVar`/`priority`。
2. 在环境变量中填入 Key（**不要写进代码或文件**）：
   - 临时（仅当前终端）：Windows `set DEEPSEEK_API_KEY=sk-xxxx`；Linux/macOS `export DEEPSEEK_API_KEY=sk-xxxx`
   - **持久（推荐）**：Windows 用 PowerShell 写入用户作用域，重启终端后仍有效、且不进 git：
     ```powershell
     [Environment]::SetEnvironmentVariable("DEEPSEEK_API_KEY", "sk-xxxx", "User")
     ```
   - 本机已按上述方式把 `DEEPSEEK_API_KEY` 配到用户环境变量，直接 `chat` 即可用；其余通道（SiliconFlow / Moonshot）按需补对应变量即可。
3. 至少启用一个供应商并配置其 Key 后即可运行：
   ```powershell
   dotnet run --project src/ThesisEngine.Cli -- chat 一级标题用几号字
   ```

`openai-compatible` 这一项即 issue 要求的 **OPEN 免费接口**：把 `baseUrl` 和 `model` 改成任意 OpenAI 兼容服务的地址与模型名，把 `apiKeyEnvVar` 改成对应的环境变量名，并设 `enabled: true` 即可。

## 5. 多模型与降级

- 运行时按 `priority` 升序依次尝试已启用且能取到 Key 的供应商。
- 主通道调用失败（网络错误、超时、返回非 2xx、JSON 解析失败、缺 Key）会**自动降级**到接下来的 2 个免费通道，最多尝试 3 个，并**逐层上报错误（共 3 层）**。
- 全部失败时返回包含 3 层失败原因的中文提示，并建议检查网络/Key 或改用 `--local`，**不会崩溃**。

## 6. 配不上怎么排查

| 现象 | 原因与处理 |
| --- | --- |
| `配置错误：未检测到可用的模型 API Key…` | 没有启用任何供应商，或启用的供应商对应的环境变量未设置/为空。请 `set`/`export` 对应变量后重试；或加 `--local` 走离线模式 |
| `配置错误：找不到聊天模型配置文件…` | 用 `--config <路径>` 显式指定，或把 `chat.models.json` 放到程序目录/当前目录 |
| 返回“所有模型通道均不可用，已按以下层级依次失败…” | 3 个通道都失败了，消息里逐条列出每层原因（通常是网络不通或 Key 无效）。逐项检查 `baseUrl`、网络、Key |
| 返回“供应商‘X’缺少 API Key…” | 该供应商启用了但对应环境变量为空 |
| 回答被拒绝“我只能回答…格式问题” | 提问不在格式范围内，属预期行为（范围硬约束） |

## 7. 相关代码位置（仅可在以下目录修改）

- `src/ThesisEngine.Core/ChatConfiguration.cs`：配置模型与加载
- `src/ThesisEngine.Core/FormatKnowledge.cs`、`ChatScopeGuard.cs`：格式知识与范围硬约束
- `src/ThesisEngine.Core/CitationConstraint.cs`：引用/标注约束
- `src/ThesisEngine.Core/OpenAiCompatibleChatService.cs`、`ChatModelRouter.cs`：模型服务与降级
- `src/ThesisEngine.Core/FormatConstraintChatService.cs`：离线规则服务
- `src/ThesisEngine.Cli/ChatCommand.cs`、`Program.cs`：CLI 接线与 `--help`
- `src/ThesisEngine.Cli/chat.models.json`：默认配置

## 8. 实测验证（真实模型通道，已跑通）

> 以下均为本机用**真实免费 Key** 运行 `chat` 的真实输出。**Key 仅存在于用户环境变量 `DEEPSEEK_API_KEY`，未写入任何代码、配置或文档文件**，可安全提交。

接入配置：启用 `deepseek` 供应商（`base_url=https://api.deepseek.com/v1`，`model=deepseek-chat`）。接口鉴权返回 **HTTP 200**，模型实际落到 `deepseek-flash`。

实测样例（节选，直接截图或复制进 issue/PR）：

```
# 主标题（区分：论文主标题 = 黑体三号居中）
$ ThesisEngine.Cli chat 论文主标题用什么字体和对齐？
你问：论文主标题用什么字体和对齐？
AI 回答：论文主标题用黑体三号、居中对齐，段前段后 0 磅、单倍行距。

# 一级 / 章节标题（黑体小四、左对齐、固定 20 磅）
$ ThesisEngine.Cli chat 一级标题用几号字、如何对齐？
你问：一级标题用几号字、如何对齐？
AI 回答：一级标题（如“一、”）用黑体小四号，左对齐；段前段后 0 行，固定行距 20 磅。

# 摘要 / 关键词（知网 / 维普 / 万方免费检测）
$ ThesisEngine.Cli chat 摘要和关键词有什么格式要求？
AI 回答：摘要与关键词的格式要求如下（仅就格式而言）：
1. 摘要正文：宋体小四、西文 Times New Roman、两端对齐、首行缩进 2 字符、1.5 倍行距。
2. “摘要”二字：黑体小四号、左对齐、固定行距 20 磅。
3. 关键词：各关键词之间用分号隔开，末尾不加标点。
4. 检测渠道：可用知网 / 维普 / 万方免费检测，不必额外花钱。

# 参考文献（GB/T 7714）
$ ThesisEngine.Cli chat 参考文献的格式规范是什么？
AI 回答：参考文献须符合 GB/T 7714；正文引用以上标[1]形式标注。

# 随意问题被代码拦截（线上路径也不调模型）
$ ThesisEngine.Cli chat 帮我写一首关于春天的诗
AI 回答：我只能回答南开大学研究生学位论文的【格式】问题……请换一个格式相关问题。
```

## 9. 关于「接口认证」

- **接口鉴权已验证**：用真实 Key 调用 DeepSeek `/v1/chat/completions` 返回 HTTP 200，证明接口打通、鉴权有效、降级链路工作正常。
- **平台级企业 / 实名认证**：DeepSeek（以及 SiliconFlow、Moonshot）控制台里的“企业认证 / 实名认证”需在对应官网用账号登录后**自助完成**，无法由本工具或命令行代替。请登录各供应商控制台按指引完成；认证通过后才会有稳定配额。
