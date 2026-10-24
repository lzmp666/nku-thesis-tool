# 字体/字号检测测试设计（Issue #10）

## 范围和状态

本文件最初交付于 Issue #10，保留其原始测试设计背景。当前实现状态以 Core 与自动化测试为准：可解析的直接格式差异会生成待用户确认的修复预览；未知值仍为 manual。以下规则输入仍是合成值，不代表学校正式要求。

所有字体名和字号均为合成测试输入，不是南开校规。规则来源未审核时不得启用正式检测或宣称论文合规。正常场景只表示该项合成断言匹配。学校来源、脚本选择与主题解析仍需负责人审核。

## 公共输入与断言约定

- 每个 FS 场景是独立 synthetic 文档；正文含一个段落，除特别说明外含一个 run。最小包包含内容类型、根关系和 word/document.xml；继承案例增加 styles.xml 及对应关系，主题案例增加 theme/theme1.xml 及对应关系。生成器和 Word 版本在实现阶段记录；目前不提供真实论文。
- w 命名空间为 http://schemas.openxmlformats.org/wordprocessingml/2006/main。位置 L1 = part word/document.xml，xpath /w:document/w:body/w:p[1]/w:r[1]；L2 为同段 w:r[2]。位置标识以文档为边界，不能跨样本复用。诊断定位到受影响 run，来源可另记录 styles/theme 部件，不能修改共享样式影响其他节点。
- 合成规则标签 F：font/name = TestFontA；S：font_size/pt = 12；H：font_size/pt = 10.5。每场景只启用明确列出的标签。标签不是正式 ruleId，不进入规则库。转为可执行测试时绑定独立测试规则及真实 ruleId，不冒用学校批准状态或伪造 source。
- 字体案例以单一脚本隔离选择：中文看 eastAsia，ASCII 英文看 ascii；相应有效字体必须解析后比较。两种脚本/主题的规则选择若当前 schema 无法表达，先提 RFC，保持设计待实现，不增设未定义字段。
- w:sz 的值是半点数；24 = 12pt、21 = 10.5pt。有效值必须从直接属性或明确的继承/主题链解析；缺失直接属性不等于最终值未知。不得根据邻近文字猜测合规。
- N：diagnostics=[]、patches=[]，仅该项已覆盖的合成断言匹配；fixLevel/severity 不适用。D：一条不匹配诊断，必须包含绑定的 ruleId、severity、location、expected、actual、fixLevel、message。U：值无法确定或能力不支持，报告未覆盖/无法判断及原因，不宣称匹配，patches=[]；若用 Diagnostic 表达，则 severity=warning、fixLevel=manual，expected 为目标、actual 为未解析状态（具体编码经 RFC/实现确定）。
- 可解析的确定不匹配使用 severity=warning、fixLevel=auto，以便生成预览；插件仍必须先展示前后值，并等待用户确认后才应用。未知项使用 manual。该状态仅表示演示修复引擎支持该属性，不代表学校规则已批准。

## 场景矩阵

| ID | 启用目标、输入与上下文 | 预期诊断/有效值 | 位置 | severity | fixLevel | 待审核来源 |
| --- | --- | --- | --- | --- | --- | --- |
| FS-001 | F；文本“中文”；直接 eastAsia=TestFontA | N；actual=TestFontA | L1 | 不适用 | 不适用 | R1/R3/R4 |
| FS-002 | F；文本“中文”；直接 eastAsia=TestFontB | D；expected=TestFontA，actual=TestFontB，字体不匹配 | L1 | warning | auto | R1/R3/R4 |
| FS-003 | F；文本“ABC”；直接 ascii=TestFontA | N；actual=TestFontA | L1 | 不适用 | 不适用 | R1/R3/R4 |
| FS-004 | F；文本“ABC”；直接 ascii=TestFontB | D；expected=TestFontA，actual=TestFontB，字体不匹配 | L1 | warning | auto | R1/R3/R4 |
| FS-005 | S；文本“中文”；直接 sz=24 | N；actual=12pt | L1 | 不适用 | 不适用 | R1/R2/R4 |
| FS-006 | S；文本“中文”；直接 sz=22 | D；expected=12pt，actual=11pt，字号不匹配 | L1 | warning | auto | R1/R2/R4 |
| FS-007 | H；文本“ABC”；直接 sz=21 | N；actual=10.5pt，不因半点误报 | L1 | 不适用 | 不适用 | R1/R2/R4 |
| FS-008 | S；文本“ABC”；直接 sz=21 | D；expected=12pt，actual=10.5pt；原因是目标不匹配而非半点无效 | L1 | warning | auto | R1/R2/R4 |
| FS-009 | S；文本“中文”；直接 sz=23 | D；expected=12pt，actual=11.5pt，目标下方半点边界 | L1 | warning | auto | R1/R2/R4 |
| FS-010 | S；文本“中文”；直接 sz=25 | D；expected=12pt，actual=12.5pt，目标上方半点边界 | L1 | warning | auto | R1/R2/R4 |
| FS-011 | S；段落样式 Child basedOn Parent；Parent rPr sz=24；Child、run 无 sz；无其他格式来源 | N；从已知父样式解析 actual=12pt；不按直接属性缺失报错 | L1（来源 styles.xml/Parent） | 不适用 | 不适用 | R1/R2/R4 |
| FS-012 | S；同 FS-011，但 run 直接 sz=22 | D；expected=12pt，actual=11pt；直接覆盖优先 | L1 | warning | auto | R1/R2/R4 |
| FS-013 | F；中文；eastAsiaTheme=minorEastAsia；主题 minorFont 对 Hans 的映射=TestFontA；语言 zh-CN；无冲突属性 | 设计目标 N，actual=TestFontA；未实现主题/脚本解析时必须 U，不得靠主题标记判匹配 | L1（来源 theme/theme1.xml） | N 不适用；U warning | N 不适用；U manual | R1/R3/R4 |
| FS-014 | F；同 FS-013，但主题部件缺失，且无可解析回退 | U；字体无法解析；不能匹配或自动替换 | L1 | warning | manual | R1/R3/R4 |
| FS-015 | S；run 无 sz；styles docDefaults/rPrDefault/rPr sz=24；无样式和直接覆盖 | N；actual=12pt，直接缺失但默认值已知 | L1（来源 styles.xml/docDefaults） | 不适用 | 不适用 | R1/R2/R4 |
| FS-016 | S；run、样式链和 docDefaults 都无 sz | U；字号未知，不能用零诊断声称匹配 | L1 | warning | manual | R1/R2/R4 |
| FS-017 | F；中文；直接 eastAsia=UnknownFont，目标 TestFontA | D；expected=TestFontA，actual=UnknownFont；仅判名称不匹配，不据名称推断本机未安装/字体无效 | L1 | warning | auto | R1/R3/R4 |
| FS-018 | S；同段两 run：“ABC” sz=24、“中文” sz=22 | 仅 L2 有一条 D；expected=12pt，actual=11pt；L1 无误报 | L2 | warning | auto | R1/R2/R4 |
| FS-019 | S；同段两 run：“中文” sz=24、“，”无 sz；所有继承/默认值也无 sz | L1 匹配；L2 为 U，不从相邻 run 猜为12pt；整段不能宣称全部匹配 | L2 | warning | manual | R1/R2/R4 |
| FS-020 | F 和 S；中文 run 字体来源全缺失，字号来源全缺失 | 两个属性均 U，分别记录未覆盖原因；patches=[]；不宣称匹配 | L1 | warning | manual | R1/R2/R3/R4 |
| FS-021 | S；sz=0（异常输入）；无回退 | U；无有效正数字号，报告异常；不得当12pt或写补丁。规则 expect=0 则应在规则加载阶段拒绝，不能进入比较 | L1 | warning | manual | R1/R2/R4 |

## 可追溯来源（待审核，不是学校正式规则）

- R1：[Issue #10](https://github.com/lzmp666/nku-thesis-tool/issues/10) 的场景类别和验收标准；[Engine API 0.1.0](../../../contracts/engine-api.md) 的 Diagnostic、定位、未覆盖报告和修复约束；[规则 schema 0.1.0](../../../contracts/rule-schema.yaml) 的字号正数/pt、severity、fix。仓库链接采用相对路径随分支版本保存。项目契约依据已对照，场景策略待负责人审阅。
- R2：[Microsoft FontSize](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fontsize?view=openxml-3.0.1)，sz 半点表示及未设置时的样式层级说明；页面版本 Open XML 3.0.1，查阅日期 2026-10-03。样式优先级和异常输入策略待样式调研复核。
- R3：[Microsoft RunFonts](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts?view=openxml-3.0.1)，脚本字体与主题属性；页面版本 Open XML 3.0.1，查阅日期 2026-10-03。主题、语言和脚本选择待样式调研复核，不能据此声称当前引擎已支持。
- R4：[规则指南](../../../docs/rule-guide.md)、[QA SOP](../../../docs/qa-sop.md)、[golden 模板](../../goldens/template.yaml)。学校条文标题、官方 URL、版本、页码、原文、适用范围和核验日期尚未提供，审核状态 pending。TestFontA/TestFontB 和 12/10.5pt 仅用于构造输入与对照，禁止迁入 approved 学校规则。后续规范任务应逐场景绑定真实来源；脚本/主题表达不足先提 RFC。

## 实现阶段的执行与修复验收

1. 依据矩阵生成独立最小 synthetic 包，将 F/S/H 映射为隔离的测试规则（不能冒充学校规则）。实现前先完成人工来源与契约表达审核；当前不提交伪造的 approved YAML。
2. 对 N/D/U 分别断言已覆盖匹配、不匹配诊断、未覆盖状态。D 必须精确校验 ruleId/位置/expected/actual/severity/fixLevel；U 的未知状态不能被吞掉。按 golden 模板人工核对，不能保存空列表就判合规。
3. suggest 未采纳时 patch=0，原诊断保留；manual 始终无自动补丁。用户采纳确定的建议后，补丁须绑定 expectedBefore 和 desiredAfter，另存输出，保护文本、关系、其他 run/样式/未知部件。前置值改变则停止并报告。
4. 如果以后将某个确定、可安全定位场景升为 auto，须另行审核并增加自动修复测试：仅改目标属性，二次已覆盖诊断=0、二次 patch=0；同时验证未解决 manual/suggest 项仍在报告中。不能用整个 docx 哈希证明幂等。
5. 未来实现加入现有测试项目后执行 README 的 restore/build/test。当前文档不在 xUnit 执行链中，CI 通过只能证明项目骨架回归，不代表这21个场景已经执行。

## 设计自检

21 个独立编号；字体正常/错误、字号正常/错误、半点与上下边界、继承/直接覆盖、主题可解析/未知、默认值/最终缺失、混合 run、标点未知及异常值均有输入、结果、定位、修复等级和待审核来源。只修改 Issue 允许路径；无真实论文、无正式校规声明、无引擎或契约改动。
