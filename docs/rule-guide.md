# 从学校条文到 YAML
1. 找到学校/研究生院官方发布文件，记录标题、官方 URL、版本/发布日期、页码/条款、原文、核验日期和适用范围。学院差异放 colleges/，冲突先请负责人判定优先级。
2. 在 Issue 附规则表：id、scope、property、expect/value/unit、severity、fix、source。把一句话拆成可独立检测的规则，不把经验当校规。
3. 对照 contracts/rule-schema.yaml：0.1.0 保留旧规则形状；0.2.0 可在显式 profile 下增加脚本字体槽、未知范围处理和矩阵追踪字段。font 使用字符串/name；font_size 使用正数/pt；line_spacing 使用正数/multiple 或 pt；page_margins 使用 top/bottom/left/right 四个数/cm。
4. 字体区分脚本和主题，行距模式、样式范围或学院覆盖无法表达时先提 RFC，不能临时塞未定义字段。id 全库唯一，禁止复用。未知范围必须 fail closed 为 unsupported。
5. `rules/nku-graduate/general.yaml` 是旧数据结构的 draft 整理稿：Demo 正文部分已按 2026 版核验，其余章节仍待逐条核验，整份文件尚不符合可加载规则 schema。`rules/nku-graduate/2026-body-v1.yaml` 是符合 0.2.0 schema 的独立 2026 正文 draft profile，不会自动启用或代替整理稿；待负责人审核来源、适用范围、语义与样本后，才可更改状态。
6. 每条规则至少有正确、错误、边界/不支持样本与 golden，记录修复等级。auto 必须安全定位并通过二次诊断/补丁为零验证。
不提供未经核验的宋体/字号/边距示例值。真实出处完整性与 property/unit 的跨字段匹配需人工及后续规则加载器验证，单靠 schema 无法代替。

## 分析引擎如何读取规则（#111）

`analyze` 现在有两条路径：`--demo` 保留旧的 3 条演示规则；不带 `--demo` 时读取 YAML 规则，默认 profile 为 `nku-graduate`。

```powershell
# 使用 rules/nku-graduate/general.yaml（draft）
dotnet run --project src/ThesisEngine.Cli -- analyze tests/fixtures/synthetic/analysis_compliant.docx

# 显式指定 2026 正文 profile
dotnet run --project src/ThesisEngine.Cli -- analyze <file.docx> --profile nku-graduate-2026-body

# 显式指定任意 0.2.0 / legacy 规则文件
dotnet run --project src/ThesisEngine.Cli -- analyze <file.docx> --rule-file path/to/rules.yaml

# 机器可读输出与标注副本
dotnet run --project src/ThesisEngine.Cli -- analyze <file.docx> --json
dotnet run --project src/ThesisEngine.Cli -- analyze <file.docx> --annotated-output artifacts/review.docx
```

加载器支持两种形状：

- legacy：`rules.<group>.<path>`，例如 `rules.body_text.font.cn`、`rules.headings.chapter.size_pt`。`general.yaml` 保持原有嵌套结构，只在每个组补 `id` 和 `meta.analysis` 开关。
- 0.2.0：`schema_version: "0.2.0"` + `rules:` 数组，使用 `id`、`matrix_id`、`scope`、`property`、`expect`、`source.locator`。

### 规则 ID、条款和 YAML 位置

每条问题都带三样可追踪信息：

- 规则 ID：legacy 使用 `R11/font.cn` 这类「组 ID + 字段路径」；0.2.0 使用 `NKU-REQ-042`。
- 条款号：`source.clause`（legacy）或 `source.locator`（0.2.0），输出时统一加 `D2026` 前缀。
- YAML 位置：`rules.body_text.font.cn`、`rules[0]` 等，能直接回到规则文件。

`general.yaml` 的 `meta.status: draft` 会被加载器读取：报告头部显示 `draft（未核验）`，置信度上限压到 0.75；它不是装饰字段，也不会被当成 approved。

### 三态与覆盖表

分析器只产生三种结果：

| 状态 | 含义 |
| --- | --- |
| 符合 | 已定位到对象，期望值/实际值均可解析且一致 |
| 不符合 | 已定位到对象，期望值/实际值均可解析且不一致 |
| 未检查 | 规则未实现、对象未识别、样式链/值无法解析，或按预警开关不逐条报告 |

关键约束：**缺失值、未知样式、未识别部件一律进入「未检查」，不会变成「符合」。** 合规样本必须同时满足「不符合=0」和「未检查非空」；后者证明没有把该测的项静默跳过。

覆盖表由加载器 + 显式 `LegacyRuleCatalog` 生成。`general.yaml` 中每个规则组、每个已映射字段都有且只有一行：

- `已实现`：比较器存在，覆盖表显示适用对象数和问题数。
- `未检查`：比较器或对象识别未实现，必须写原因，例如页眉未读取（T5）、目录条目对象待定（T6）、图表由 #92 的 `FigureTable*` 处理、封面固定版式（RFC-6）。

同一规则、同一期望/实际值的重复段落会合并为一条问题，`occurrences` 保留每个段落的位置；文本报告列出前几处，JSON 保留全部。这样既能给出完整问题列表，又不会把几十个同值段落刷成几十条。

### 预警开关

`general.yaml` 的 `meta.analysis.report` 控制报告层：

```yaml
meta:
  analysis:
    report:
      default: true
      suppress: []      # 例： [toc.entries] 只计数，不逐条列出
```

组级也可写 `report: false`。关闭后仍会执行比较并计入覆盖表的 `问题` 数，但不会生成逐条问题；未检查清单会保留「按预警开关未逐条报告（N 处不符合）」。这是报告开关，不是规则值开关，不会把不符合改成符合。

### 对象识别

对象识别复用 `ThesisStructureClassifier` 的 `Section` / `Kind` / `HeadingLevel`，不另写一套章节识别。映射关系：

| 结构决策 | 分析对象 |
| --- | --- |
| Body + Heading(1/2/3/4) | 章标题 / 一级 / 二级 / 三级标题 |
| Body + Paragraph | 正文段落 |
| ChineseAbstract / EnglishAbstract + Heading | 中/英文摘要标题 |
| ChineseAbstract / EnglishAbstract + Paragraph/Keywords | 中/英文摘要正文（含关键词标签） |
| Contents + Heading / ContentsEntry | 目录标题 / 目录条目 |
| References + Paragraph | 参考文献正文 |
| Appendix + Paragraph | 附录正文 |
| Acknowledgments + Paragraph | 致谢正文 |
| FrontMatter | 封面/题名页（v1 不检查） |

参考文献、附录、致谢的标题映射为「章标题」并使用 R7 规则；对应组的 `heading` 覆盖行标注为「复用 R7」，不重复产生一套检查。

### 当前覆盖边界

| 已实现 | 未检查（原因） |
| --- | --- |
| 节纸张 A4、四边页边距（R1） | 页眉/页脚距边界、装订线（T5） |
| 正文中/西文字体、字号、对齐、首行缩进、行距、段前段后（R11） | 页眉、页码（主文档读取器不读页眉页脚） |
| 章/一/二/三级标题字体、字号、加粗、行距、段前段后（R7–R10） | 章/一级标题对齐模式（A2）、模式二缩进（A2） |
| 中/英文摘要标题与正文、关键词标签加粗（R4–R5） | 目录条目（T6） |
| 目录标题（R6） | 图题/表题、公式（#92） |
| 参考文献、附录、致谢正文（R14/R15/R17） | 符号注释表/个人简历（对象识别未实现） |
|  | 封面/题名页（RFC-6） |

### general.yaml 规则组覆盖表（116 个可追踪槽位）

| 规则组 | 条款 | 已实现 | 未检查 |
| --- | --- | --- | --- |
| R1 页面设置 | 4.1 表4.1 | 2 | 2（页眉页脚距、装订线 T5） |
| R2 篇眉 | 4.1 表4.1；3.4 | 0 | 4（主文档读取器不读页眉） |
| R3 页码 | 4.1 表4.1；3.4 | 0 | 4（页码字段未读取） |
| R4 中文摘要 | 4.5 表4.4 | 13 | 0 |
| R5 英文摘要 | 4.5 表4.4 | 13 | 0 |
| R6 目录 | 4.6 表4.5 | 7 | 1（条目 T6） |
| R7 章标题 | 4.7 表4.6 | 6 | 1（对齐 A2） |
| R8 一级标题 | 4.7 表4.6 | 6 | 1（对齐 A2） |
| R9 二级标题 | 4.7 表4.6 | 7 | 1（模式二缩进 A2） |
| R10 三级标题 | 4.7 表4.6 | 7 | 1（模式二缩进 A2） |
| R11 正文 | 4.7 表4.6 | 8 | 0 |
| R12 图题/表题 | 4.7 表4.6 | 0 | 1（#92） |
| R13 公式 | 4.7 表4.6 | 0 | 1（#92） |
| R14 参考文献 | 4.8 表4.7 | 6 | 1（标题复用 R7） |
| R15 附录 | 4.8 表4.7 | 8 | 1（标题复用 R7） |
| R16 符号表/简历 | 4.8 表4.7 | 0 | 7（对象识别未实现） |
| R17 致谢 | 4.8 表4.7 | 5 | 1（标题复用 R7） |
| R18 封面 | 4.2/4.4 | 0 | 1（RFC-6） |

每个槽位的实际状态以 `analyze` 输出的「规则覆盖表」或 `--json` 的 `coverage[]` 为准；上表用于快速核对没有整组静默跳过。

### 回归样本

- `tests/fixtures/synthetic/analysis_compliant.docx`：A4、3.8/3.2 cm 页边距、宋体/Times New Roman 12 pt、两端对齐、首行缩进 2 字符、固定值 20 pt；预期 0 条不符合、未检查非空。
- `tests/fixtures/synthetic/analysis_noncompliant.docx`：同一文档的第一段故意改为黑体/Arial 10.5 pt、左对齐、0 缩进、单倍行距、段前段后 6 pt；预期 8 条不符合。
- `tests/fixtures/synthetic/full-thesis/synthetic_complete_thesis.docx`：完整合成论文，用于验证问题合并、覆盖表、JSON 和确定性。
- 生成脚本：`tests/fixtures/synthetic/generate_analysis_samples.py`。

这些样本是合成数据，不代表真实论文；规则仍为 draft，报告不能作为正式合规结论。
