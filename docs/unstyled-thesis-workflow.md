# 无样式论文：结构识别、规范匹配与原文排版

实现日期：2026-10-26。1.5.4 预览版将角色确认和整篇排版接入 Word 功能区，同时保留 CLI。配置仍为部分覆盖的南开 2026 草案，不代表所有论文格式均已覆盖。

## 在 Word 中使用

打开 DOCX 后，进入“论文格式检查 → 大纲确认与排版”。窗口展示当前内存文档中的各段原文、所属部分、类型、标题层级和判断依据，可点击定位核对。

所有行初始均未确认。可明确点击“确认全篇高置信规则项”，再按中文摘要、英文摘要、参考文献、附录等部分筛选，核对并确认不确定项；也可逐行修改角色和标题层级。角色被编辑后需要重新勾选确认。未知角色、表格内部、复杂内联结构或不安全锚点不能直接批准。

选择模式一/模式二，勾选已核对草案范围后，点击确认并排版。**只有本轮明确勾选的块会进入写入计划**；未勾选、未覆盖部分保留原格式，计入结果说明。程序在当前 Word 文档应用字体、字号、粗体及目标段落格式，不重开或自动保存文档。

写入前检查当前文档快照、段落与每个字符跨度；写入后从 Word 重新生成快照，由引擎核对全篇原文和目标格式。排版分片运行，回到消息循环时没有悬挂的 Word 撤销事务。检查不通过时恢复本轮已完成格式，后续编辑发生冲突时停止并说明。

“撤销整篇排版”只恢复本轮涉及的格式属性。原文或这些属性被再次编辑时停止，避免覆盖作者后续更改。文档保存仍由作者决定。

本机完整合成稿验收见 [2026-10-26 记录](demo/word-outline-acceptance-20261026.md)：286 项回归通过；Word 中明确勾选的 83 段回读通过、分析剩余0项，整篇撤销恢复原有10组问题。一次对比中，UI STA 调度和增量重绘将写入、回读与重分析从95.1秒降至19.1秒；这是单篇单次测量，不是性能上限或真实论文准确率。

## 处理链路

1. 从原生 DOCX 提取段落、表格归属、题注、标题样式和不支持结构；保留所有原始文字及 part/结构路径。
2. 建立 `DocumentSourceMap`：段落 ID、全文哈希、相邻段上下文和 UTF-16 run 跨度。改字体造成 run 拆合时可重新投射文字跨度；重复内容或内容变化不能靠“第几个同名句子”猜位置。
3. 在没有标题样式时，组合中英章节词、编号连续性、目录范围、正文衔接和文末边界进行结构分类。分数是证据强度，并非正确率。闭合证据不足时保留复核状态。
4. 可选本机 Ollama 提议少量不确定块的角色。模型仅接收已有 BlockId 和有界上下文；未知/重复编号、过期快照、非法枚举、额外属性或虚构证据会被拒绝。候选不会直接生成修复补丁。
5. 作者一次确认大纲，之后将角色匹配到明确选择的规范配置，统一处理可覆盖部分。规范值来自版本化配置，模型不提供字体、字号、行距或写入坐标。
6. CLI 另存一个明确指定的新文件，验证输入快照、原文文字、源锚点、目标格式与其他包部件；Word 则在当前文档内写入并回读，不自动保存。

## 使用

只分析结构，不调用模型、不改 DOCX：

```powershell
dotnet run --project src/ThesisEngine.Cli -c Release -- structure article.docx --report structure.json
```

报告包含每块角色、证据、置信度、源位置和 `approvalTemplate`。将待复核项确认后的模板另存为 `outline-approved.json`，再运行：

```powershell
dotnet run --project src/ThesisEngine.Cli -c Release -- structure article.docx `
  --approved-outline outline-approved.json --nku2026-draft mode1 `
  --format-output article-formatted.docx --report formatted-report.json
```

`mode1` 是阿拉伯数字层级标题；`mode2` 是另一种章/节模式。当前不自动改写编号或题名，不替作者选择模式。标题会清除与目标冲突的旧缩进，不能仅改字体后遗留正文首行缩进。

`--nku2026-draft` 明确表示作者选择的预览配置。数值已与官方 2026 DOCX 第3.3节、表4.4–4.7独立核对，但配置仍为 Draft，覆盖范围不完整。未知结构、目录条目、封面、页眉页码、表格单元格、复杂图表和公式会记录在 `Omissions`，不会把它们默认为已合规。关键词标签局部加粗、题注移动/编号间距也尚未覆盖。

## 可选 AI

先由操作者配置本机 Ollama 和可用模型，并在服务端设置 `OLLAMA_NO_CLOUD=1` 或 `disable_ollama_cloud=true` 后重启。只有 loopback 地址不能证明模型实际在本机计算。模型权重不由插件自动下载，本项目也未训练或实测某个具体模型。

```powershell
dotnet run --project src/ThesisEngine.Cli -c Release -- structure article.docx `
  --ollama-model <已配置的本机模型名> --local-model-confirmed --report ai-outline.json
```

默认连接 `http://127.0.0.1:11434/api/generate`，关闭 HTTP 代理和重定向，拒绝显式 `:cloud` 模型名；每次请求使用 JSON schema、覆盖响应体读取的期限和大小上限。默认最多四批、每批十六个待确认块，并缓存已验证的同快照响应。达到上限、超时或模型不可用时保留本地结果及明确警告。

模型可能被文章中的恶意指令影响而产生错误候选；本实现依靠候选隔离、原文证据与作者确认控制写入，不宣称模型语义判断百分之百可靠。参考：[Ollama structured outputs](https://docs.ollama.com/capabilities/structured-outputs)、[云功能关闭方式](https://docs.ollama.com/faq#how-do-i-disable-ollamas-cloud-features)。

## 训练与评测

确认后的标签可以通过 `--training-output labels.jsonl` 导出；该选项必须同时提供 `--approved-outline`。只有作者明确确认的标签可作为训练样本，AI 候选不充当真值。每条样本有 `documentGroup`，训练/验证/测试须按整篇文档分组，不能把同篇相邻段落随机拆开制造高分。

目前已实现标注导出和可替换的 `IThesisStructureModel` 接口；尚无获授权、独立标注的真实论文测试集，因此没有真实论文准确率或已训练模型的承诺。应分别报告中文、英文、无样式和异常结构的各类 precision/recall、macro-F1、边界误判、错误修改数，以及冷启动/缓存命中延迟。

开源选型和许可证证据见 [调研文档](research/unstyled-thesis-open-source-2026-10-26.md)。本轮借鉴原生 OOXML 映射、分阶段分类及结构化模型输出，不引入 PDF 坐标作为 Word 写入坐标，也没有复制上游代码或捆绑大型模型。

完整无样式合成稿可由标准库脚本生成：

```powershell
python tests/fixtures/synthetic/full-thesis/generate_unstyled_thesis.py --output <新文件.docx>
```

该稿基于已有完整六章合成论文，去除标题样式和直接文字排版，保留中英文摘要、目录、图表、公式、参考文献和附录；不是几行测试文字。原稿曾在 Word 中显示为14页，新稿实际页数会随排版变化，应在 Word 内另行确认。
