# 从学校条文到 YAML
1. 找到学校/研究生院官方发布文件，记录标题、官方 URL、版本/发布日期、页码/条款、原文、核验日期和适用范围。学院差异放 colleges/，冲突先请负责人判定优先级。
2. 在 Issue 附规则表：id、scope、property、expect/value/unit、severity、fix、source。把一句话拆成可独立检测的规则，不把经验当校规。
3. 对照 contracts/rule-schema.yaml：0.1.0 保留旧规则形状；0.2.0 可在显式 profile 下增加脚本字体槽、未知范围处理和矩阵追踪字段。font 使用字符串/name；font_size 使用正数/pt；line_spacing 使用正数/multiple 或 pt；page_margins 使用 top/bottom/left/right 四个数/cm。
4. 字体区分脚本和主题，行距模式、样式范围或学院覆盖无法表达时先提 RFC，不能临时塞未定义字段。id 全库唯一，禁止复用。未知范围必须 fail closed 为 unsupported。
5. `rules/nku-graduate/general.yaml` 是旧数据结构的 draft 整理稿：Demo 正文部分已按 2026 版核验，其余章节仍待逐条核验，整份文件尚不符合可加载规则 schema。`rules/nku-graduate/2026-body-v1.yaml` 是符合 0.2.0 schema 的独立 2026 正文 draft profile，不会自动启用或代替整理稿；待负责人审核来源、适用范围、语义与样本后，才可更改状态。
6. 每条规则至少有正确、错误、边界/不支持样本与 golden，记录修复等级。auto 必须安全定位并通过二次诊断/补丁为零验证。
不提供未经核验的宋体/字号/边距示例值。真实出处完整性与 property/unit 的跨字段匹配需人工及后续规则加载器验证，单靠 schema 无法代替。
