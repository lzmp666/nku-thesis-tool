# Golden 人工预期记录

template.yaml 和 synthetic-example.yaml 使用 JSON 形式的 YAML 1.2，可用 JSON 或 YAML 1.2 解析器读取。它们是测试预期，不是 contracts/rule-schema.yaml 的正式规则文件。

示例提供完全自建的 XML 输入、半点字号人工推导、ruleId/location/expected/actual/severity/fixLevel、未检查范围及二次 diagnostic/patch。SYNTHETIC-SIZE-001 是隔离测试标识，绑定实际运行规则前需实现与负责人审核，不放入学校批准规则集。

执行前保持 execution.status=not_executed、review_status=pending、reviewed_by/on=null；after 的空数组是待验证目标，不是已执行结果。禁止仅因 diagnostics=[] 或 patches=[] 判断合规。审核状态为 approved 时必须填写真实执行环境、commit、日期与独立审核者，并确认每个范围实际覆盖。

suggest 未被采纳时不得执行补丁，原诊断仍存在；accepted_suggestion 记录采纳后的候选预期，不是当前可执行计划。需验证 expectedBefore、另存输出、内容/关系/样式/节保持、二次无新补丁及未解决项保留。manual 不得自动修复。after 仅针对已覆盖且已处理的属性，不代表整篇论文合规。

本次自审日期 2026-10-06：核对 JSON 可解析、输入 XML、22 半点=11pt、24 半点=12pt、定位与字段一致。当前引擎未实现检测/修复，未伪造运行记录；独立审核待完成。
