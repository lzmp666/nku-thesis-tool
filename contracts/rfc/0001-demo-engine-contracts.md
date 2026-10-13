# RFC 0001 — Demo 引擎契约（draft/demo 阶段）

- 状态：accepted
- 负责人：@lzmp666
- 对应 Issue：#33
- 对应 PR：#36
- 影响代码：`src/ThesisEngine.Core/DemoContracts.cs`
- 影响下游 PR：#35（FontSizeAnalyzer）

## 1. 背景与例子

`contracts/engine-api.md` 已经定义了正式引擎的目标接口（Load / Analyze / PlanPatches / ApplyPatches / GenerateReport），但当前代码只实现了 `DocumentLoader`。#29 要做一个最小可演示的字体/字号诊断器，需要在不污染正式合规路径的前提下，先跑通「读取文档 → 产出诊断 → E2E 对照 manifest」这条链路。

演示样本（#27 合入的 `tests/fixtures/synthetic/demo/`）里，`known_issues_sample.docx` 应有三条诊断：

| ruleId | location | expected | actual |
|---|---|---|---|
| `demo/east-asian-font` | paragraph[0]/run[0] | 宋体 | 黑体 |
| `demo/western-font` | paragraph[1]/run[0] | Times New Roman | Arial |
| `demo/font-size` | paragraph[2]/run[0] | 12pt | 10.5pt |

为了让诊断器、文档读取器、未来的 E2E 脚本共享同一份类型，需要先落一组 demo 专用契约。这些类型**只服务 demo sprint**，显式带 `Demo`/`draft` 前缀，不能直接当成正式 NKU 合规模型。

## 2. 提案

新增 `src/ThesisEngine.Core/DemoContracts.cs`，包含：

- `DemoRuleSet`：demo 规则集配置。当前字段 `RuleVersion`、`IsDraft`、`EastAsiaFontExpected`、`AsciiFontExpected`、`FontSizeExpectedPt`，内置 `NkuDraftSample` 静态实例（宋体 / Times New Roman / 12pt）。
- `DocumentSnapshot` / `ParagraphSnapshot` / `TextRunSnapshot`：只读文档中间表示，给诊断器消费。`TextRunSnapshot` 暴露 `EastAsiaFont`、`AsciiFont`、`FontSizePt`、`Text`，未解析到的字段允许为 `null`。
- `Diagnostic`：诊断条目，字段对齐 `engine-api.md` 中 Diagnostic 的形状（`RuleId`、`Severity`、`Location`、`Expected`、`Actual`、`FixLevel`、`Message`）。
- `AnalysisResult`：分析结果，含 `Diagnostics`、`RuleVersion`、`Coverage`、`UnsupportedScopes`。
- `IDemoDocumentReader` / `IDemoAnalyzer` / `IDemoAnalysisService`：读取器 / 诊断器 / 组合入口接口。

约束：
- 所有类型命名空间在 `ThesisEngine.Core`，文件级 XML 注释标注「draft/demo，不能用于正式 NKU 合规检查」。
- rule ID 统一使用斜杠风格 `demo/<rule-name>`，与 #27 `manifest.json` 对齐；本 PR 不内置任何具体规则实现（分析器在 #35）。
- 不修改 `contracts/engine-api.md`、`contracts/rule-schema.yaml`、`rules/` 下任何正式规则文件。

## 3. schema / API 影响

- `engine-api.md` 中的 `Analyze(document, approvedRuleSet) → AnalysisResult` 是正式接口；本 RFC 提供的是它在 demo 阶段的**最小投影**，字段名与正式版保持一致，避免未来正式化时重命名。
- 正式接口要求 `ApprovedRuleSet`；demo 阶段用 `DemoRuleSet`（带 `IsDraft=true`），通过命名前缀强制区分。
- `Diagnostic.Location` 在 demo 阶段使用 `<part>#paragraph[i]/run[j]` 形式；#27 manifest 中是 `paragraph[i]/run[j]`，E2E 脚本按后缀匹配，本 PR 不改动 manifest。
- 不引入任何外部依赖，纯 `record` + `interface`，net10.0 可编译。

## 4. 兼容与迁移

- 这是新增文件，不改动任何已有公开类型，对 main 上现有 `DocumentLoader`、CLI、测试零破坏。
- 未来正式化时（demo sprint 结束），迁移路径：把 `DemoRuleSet` → `ApprovedRuleSet`、`DemoRuleSet.NkuDraftSample` → 从 `rules/nku-graduate/general.yaml` 加载、`IDemoAnalyzer` → `IAnalyzer`。届时再开 RFC，不在本 PR 内做。
- 已在 `FontSizeAnalyzer`（#35）中预演了消费方式，确认字段形状够用。

## 5. 测试 / 回滚

- 编译期：`dotnet build` 通过。
- 运行期：`dotnet test` 22/22 通过（其中 FontSizeAnalyzer 8 个单测消费本 RFC 的类型）。
- 回滚：本 PR 单文件新增，直接 revert 即可，无迁移脚本、无数据、无配置。

## 6. 替代方案

- **方案 A（本 RFC）**：在 Core 里加一组 demo 专用 record/interface。
- **方案 B**：把这些类型直接塞进 `contracts/engine-api.md` 对应的正式接口里。否决：会让 draft/demo 模型污染正式合规路径，违反 #29 文件边界约定。
- **方案 C**：让 `FontSizeAnalyzer` 自己内联私有 record，不抽公共契约。否决：#27 E2E 脚本、未来第二个诊断器、读取器都需要共享形状，内联会复制并漂移。
- **方案 D**：等正式规则 schema 定稿后再写 demo 契约。否决：会卡住 #27 验收和 #29 演示，违反 sprint 节奏。

## 7. 审批

- [x] @lzmp666 review 通过并标记 accepted（2026-10-13）
- [ ] PR #36 合并入 main
- [ ] #35 rebase 到最新 main，确认 diff 只剩 `FontSizeAnalyzer.cs` + `FontSizeAnalyzerTests.cs`
