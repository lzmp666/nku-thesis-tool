# Engine API 0.1.0
此文件是目标接口契约；当前代码仅实现 DocumentLoader，检测/修复未实现。
- Load(path) → LoadedDocument：只读打开 .docx，验证必需部件，错误抛出明确异常，不写入原文件。
- Analyze(document, approvedRuleSet) → AnalysisResult：diagnostics、ruleVersion、coverage、unsupportedScopes。
- PlanPatches(result) → PatchPlan：仅 auto 可自动执行；suggest 需要用户采纳，manual 给操作说明。
- ApplyPatches(document, plan, outputPath) → ApplyResult：另存，不覆盖输入；补丁绑定部件/节点定位、expectedBefore、desiredAfter，前置条件失败停止并报告。
- GenerateReport(before, after, patchResult) → Report：列出已查、未查、跳过、失败、规则版本/出处和输出路径。
Diagnostic：ruleId、severity、location(part/xpath 或稳定节点标识)、expected、actual、fixLevel、message。
Patch：ruleId、location、expectedBefore、desiredAfter、reason。补丁不得改论文内容、丢失关系或破坏未知部件。
合规仅针对启用且实际覆盖的规则；不支持/未检查项不能因 diagnostic=0 被宣称全面合规。
完成判据：已覆盖且可修复项再次 Analyze 的 diagnostic=0，并且再次 PlanPatches 的 patch=0；manual/suggest 未解决项单独保留。
批次失败不发布半成品。需要时用规范化 OOXML 验证无关内容保持，禁止整包文件哈希作为幂等依据。
Core 不依赖 Word COM；未来 VSTO 与 net10.0 CLI/Core 通过版本化进程协议通信，避免 .NET Framework VSTO 直接引用不兼容 Core。
