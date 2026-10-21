# RFC 0002 — 2026 版研究生论文正文文字规则

- 状态：proposed
- 负责人：@lzmp666
- 对应 Issue：#74
- 来源矩阵：`docs/rule-notes/nku-2026-coverage-matrix.md`，M26-036
- 规范依据：南开大学研究生学位论文写作规范（2026版）第 4.7 节表 4.6
- 官方来源：[研究生院规范页面](https://graduate.nankai.edu.cn/2017/0222/c23238a56863/page.htm)，[页面所附原始 DOCX](https://graduate.nankai.edu.cn/_upload/article/files/46/28/7965029e47d08d27f560f8393fd8/28100848-9be2-4f1b-8f74-5899900cf936.docx)
- 本 PR 范围：仅提出规则契约与实现边界；负责人接受 RFC 前不改规则 YAML、分析器或插件协议。

## 1. 背景与例子

覆盖矩阵 M26-036 记录的正文段落要求包括中文宋体 12 pt、英文 Times New Roman 12 pt、两端对齐、首行缩进 2 个汉字、固定值行距 20 pt、段前段后 0 pt。包含数学表达式的段落可以按表达需要调整行距。

当前 `DemoRuleSet.NkuDraftSample` 只有三项演示值，读取器把主文档段落交给分析器，分析器逐 run 检查字体槽和字号；它没有可靠的正文部件分类。因此直接把 Demo 规则改名为 2026 正式规则，会把字体/字号要求套到摘要、标题、目录、参考文献或表格等未授权范围，且会把未知格式报告成错误。该风险必须先在契约中解决。

## 2. 提案

### 2.1 配置版本与状态

新增独立的 2026 正文 profile，例如 `rules/nku-graduate/2026-body-v1.yaml`，不得把规则叠加进既有 `general.yaml`，也不得改变其 draft/demo 语义。

- `profileId`：`nku-graduate-2026-body`
- `ruleVersion`：`nku-graduate-2026-body-v1`。这是仓库配置版本，不是学校规范另一个版本；任何规则值、范围或解释变化均递增 profile 版本。
- `status`：初始为 `draft`。只有规则负责人审阅来源、适用对象和测试样例后，才可改为 `approved`。
- 依据：2026 版第 4.7 节表 4.6，来源记录须保留研究生院官方页面、原始附件、条款定位、原文摘要和核验日期。版本字段始终写 `2026`，不得从网页发布日期推断规范版本。

M26-036 目前列出的 `NKU-REQ-036` 是候选号，不是正式分配。由于首批诊断拆成中文字体、英文/西文字体和字号三个可独立失败的属性，提议在全库核重后分配三个 ruleId，并在配置中都回链 M26-036。候选可用 `NKU-REQ-042`、`NKU-REQ-043`、`NKU-REQ-044`；负责人接受 RFC 时须明确批准或重编号，未批准前不得写入正式规则文件。

### 2.2 首批规则值

| 候选规则 | 属性 | 期望值 | OOXML 解释 |
|---|---|---|---|
| `NKU-REQ-042` | 中文文字字体 | 宋体 | 对含汉字的正文片段核验有效 `eastAsia` 字体槽 |
| `NKU-REQ-043` | 英文文字字体 | Times New Roman | 对含拉丁字母的正文片段核验有效 `ascii` 与 `hAnsi` 字体槽 |
| `NKU-REQ-044` | 正文字号 | 12 pt | 对正文可见文字核验有效字号；OOXML `w:sz` 半磅值 24 对应 12 pt |

上述值只对应表 4.6 中的正文段落文字。两端对齐、首行缩进、行距和段前段后列为后续批次，不因本 RFC 而宣称已检查。含数学表达式时，表 4.6 明确允许调整的是该段行距；本批不检查行距，因此该例外既不豁免同段普通文字的字体/字号，也不为公式对象推断字体/字号。

### 2.3 适用范围与保守分类

规则执行前必须由范围分类器把内容标记为 `confirmed-body-prose`、`excluded` 或 `unknown`。只有明确判定为 `confirmed-body-prose` 的可见正文文字可应用这三项规则；未知范围进入 `unsupportedScopes`，不报告为符合或不符合。

纳入范围：

- 由用户明确选择的研究生 2026 版 profile 下，论文正文中已确认的普通段落文字。
- 中英文混排的同一 run 可以分别检查两个适用字体槽；字号对其中的普通文字统一检查。
- 直接格式、字符样式、段落样式及其 `basedOn` 链、文档默认值可解析时，按有效格式检查；主题字体必须能确定解析为具体字体名。

排除范围：

- 封面、题名页、声明/授权书、中英文摘要、目录及图表目录、正文标题、参考文献、附录、索引、致谢、简历/成果。
- 表格单元格、脚注/尾注、页眉页脚、文本框/浮动对象、图片内文字、非正文部件、字段指令、非文字公式对象。
- 数字/标点/符号独立片段不单独推断其应使用的字体脚本槽；仅在同一普通文字片段含汉字或拉丁字母时，对相应槽位检查。中西文混排按各自可识别字母分别核验。

章节边界不能可靠识别、标题/目录字段分类不确定、样式链循环或缺失、主题字体无法解析、字体槽没有可确定的有效值时，相关内容为 `unknown`。不得用操作系统字体回退、`ascii` 值推断 `eastAsia` 值，或仅靠“某段看起来像正文”把它提升为已确认范围。实现应在诊断/覆盖结果中指出具体未覆盖位置或范围。

文档类型/规范 profile 不得仅凭默认值或推测自动选定。若学术/专业学位及硕士/博士各版本的适用范围还未由负责人确认，界面或调用方须显式选择 profile，并提示其适用范围待确认；不能宣称该配置适用于全部南开学位论文。

范围分类器至少使用明确的论文部件边界、已识别的章节标题/大纲级别和 Word 内容结构；仅因内容位于 `word/document.xml` 就不能标为正文。表格 ancestry、目录域、非正文部件应直接排除。若正文起止点、标题身份或摘要/参考文献等部件边界无法可靠确定，整个受影响片段维持 `unknown`；不得静默扩大到“全文所有段落”。

### 2.4 诊断与修复级别

- 对已确认范围内、能确定的实际值与期望值不一致项，输出现有 `Diagnostic` 字段：正式候选 ruleId、severity、位置、期望值、实际值、依据/条款与解释性中文提示；`AnalysisResult.ruleVersion` 返回 profile 版本。
- 实际值无法解析时，列为未覆盖/需人工确认，不伪造 `actual`，不算通过，也不伪报格式错误。
- 在本批中三项均只诊断，不自动修复。候选配置的 `fix` 使用 `manual`；任何自动修复须另行证明范围与格式保全安全并单独审定。
- `unsupportedScopes` 记录排除/未知内容的类别和原因；覆盖率只针对被确认并实际检查的范围计算，不能把排除内容计入“已检查”。

### 2.5 规则 schema / API 影响

现有 `contracts/rule-schema.yaml` 的 `scope` 只能表达 `body`、`heading`、`section`，`font` 也没有字符脚本槽位字段；它不能表达上面的 profile 版本、正文范围确认状态、按脚本选择字体槽或未知值策略。

RFC 接受后，在实现 PR 中将 schema 升至 `0.2.0`，以向后兼容的可选字段表达：

- profile 元数据：`profile_id`、`rule_version`、学校规范版本、状态及来源；
- 明确范围策略：纳入对象、排除对象、分类不确定时的行为；
- 字体规则的 `text_script` / `font_slots`，至少能表达 `eastAsia` 与 `ascii+hAnsi`；
- 对未解析值的处理策略，以及 M26 矩阵回链。

配置形状示意（字段名仍由接受 RFC 后的 schema 定稿；此处不是可加载规则文件）：

```yaml
schema_version: "0.2.0"
profile:
  profile_id: nku-graduate-2026-body
  rule_version: nku-graduate-2026-body-v1
  status: draft
  standard_version: "2026"
  selection: explicit
  scope_policy:
    include: [confirmed_body_prose]
    exclude: [heading, abstract, toc, references, table_cell, footnote, header_footer, text_box, equation_object]
    unknown: unsupported
rules:
  - id: NKU-REQ-042
    matrix_id: M26-036
    property: font
    text_script: east_asian
    font_slots: [eastAsia]
    expect: {value: 宋体, unit: name}
    severity: warning
    fix: manual
```

已有 `Diagnostic`、`AnalysisResult` 形状可承载规则版本、期望/实际值和未覆盖范围，不提议新增 Word 插件协议字段。本实现须在 Core 内部增加显式范围分类数据，不得由 `FontSizeAnalyzer` 继续无条件遍历所有 paragraph/run。schema 升级及配置加载失败时应 fail closed，不回退到 Demo profile。

### 2.6 实现验收用例

接受 RFC 后的实现 PR 至少增加以下合成样例与 golden：

1. 正文中文/英文分别符合字体与字号时零诊断；各自字体错误和字号错误可独立定位。
2. 同一个 run 混排中英文时分别使用 `eastAsia` 与 `ascii+hAnsi`；只有实际出现对应字母类型才检查对应字体槽。
3. 继承自段落样式、字符样式、`basedOn` 链和文档默认值的正确/错误值都按有效格式判断。
4. 缺失字体值、未解析主题字体及不确定正文边界报告 unsupported，不报告违规或通过。
5. 标题、摘要、目录、参考文献、表格、脚注、页眉和公式对象不套正文规则；中间出现无法分类的结构时显式保留未覆盖信息。
6. 含公式的正文段落：普通文字字体/字号仍适用；本批不对段落行距或公式对象作诊断。
7. `ruleVersion`、来源条款、中文消息、fix level 与矩阵回链均在 golden 中断言；既有 `demo-*` fixtures 和 Demo 结果保持原样。

## 3. 兼容与迁移

接受后先发布仍为 `draft` 的独立 profile 与 schema 0.2.0；Demo profile、`demo/*` ruleId、CLI/VSTO 行为保持独立。新分析路径必须显式收到 profile，不得让旧调用默认升级到 2026 profile。负责人批准规则状态前，插件需将报告标为草案/试用并显示实际覆盖范围；拒绝或回滚时删除新 profile 和分类器调用，既有 Demo 行为不变。

若旧的规则加载器无法识别 schema 0.2.0，须明确报不兼容并停止本 profile 加载，不允许忽略新字段后错误扩大 scope。若 implementation review 决定 schema 采用不同的兼容策略，需先更新本 RFC 并由负责人重新接受。

## 4. 测试 / 回滚

- RFC 阶段只校对条文、范围表和字段示例；本 PR 不引入生产代码或改变规则输出。
- 实现阶段按 §2.6 运行新增 Core 单测和 synthetic E2E，并运行现有 Demo 回归、Release 构建和 CI。
- 每条 golden 断言规则版本、依据、规则 ID、位置、期望/实际值、处理分类和 `unsupportedScopes`。
- 回滚通过 revert 实现 PR；移除 2026 profile 与新增 schema 版本路由，不影响 Demo 和已有报告格式。

## 5. 替代方案

- **直接复用 Demo 规则**：拒绝。Demo 当前按所有可读段落遍历，规则 ID、版本和范围都不表达 2026 正式来源，可能误套摘要、标题、目录、参考文献及表格。
- **先把全部表 4.6 规则一次性加入 YAML**：拒绝。对齐/缩进/行距等需段落分类和 Word 有效布局解释；数学表达式行距有例外，且这些规则尚无绑定样例。
- **对无法识别的值按不符合或按合规默认值处理**：拒绝。两种做法都会把未知事实误报成结论；应报告未覆盖/需人工确认。
- **全篇默认使用研究生 2026 profile**：拒绝。摘要、标题、目录、参考文献和表格有不同条款，论文类别适用范围也未完全确认。

## 6. 负责人决议

- [ ] 接受本 RFC 与三项 ruleId 候选号。
- [ ] 拒绝；说明原因：
- [ ] 接受但要求修改；说明具体字段/范围：

在负责人接受前，本任务不进入 schema、规则 YAML、Core 分析器或 Word 插件实现阶段。
