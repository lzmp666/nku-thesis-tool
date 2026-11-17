# 南开大学研究生学位论文格式检查与整改工具

面向论文排版检查与辅助整改的开源原型。产品交互界面运行在桌面版 Word 的 VSTO 任务窗格中；.NET 核心引擎和 CLI 用于分析、自动化测试及演示样例。当前只覆盖少量合成格式规则，不是南开大学正式规范的完整实现，也不代表论文合规结论。

> 当前仓库已公开。演示请使用仓库中的合成样例；不要提交真实论文、个人信息或未经授权的文档。

第一次进来（人或 AI）先看[十分钟上手](docs/ai-onboarding.md)；动手改代码前读 [AGENTS.md](AGENTS.md)，那里写了不能碰的目录和第一原则。

## 当前进度

| 能力 | 状态 |
| --- | --- |
| .NET 10 核心引擎与 CLI | 可运行；CLI 共 21 个子命令（analyze、analyze-styles、analyze-figures、analyze-formula、analyze-document、analyze-footnotes、analyze-template、analyze-abstract、analyze-back-matter、analyze-body-text、analyze-headings、analyze-toc、report、apply-nku-format、apply-figure-format、apply-formula-format、apply-document-format、full-format、undo、chat、structure）；`analyze --citations` 另提供参考文献一致性检查模式 |
| 无样式论文结构与角色排版 | Word 1.5.4 预览：大纲确认、明确勾选范围、原位整篇排版、写后回读与整篇撤销；核心/CLI 保留可选 Ollama 接口 |
| 中文诊断报告 | 可运行 |
| 将诊断写入新的 Word 文档副本（批注和黄色高亮） | 可运行；已在 Microsoft Word 桌面版打开验证 |
| Word 插件前端（中文任务窗格、原文定位、锚点联动、单项/批量修复预览） | 已合入 main（#65、#69）；已有合成文档实测记录，完整阶段验收仍由 #51、#63、#52 跟踪 |
| 已修复的误判 | 三线表修复后复检误报「有完整内边框」已修复（#106）；full-format 二次运行幂等、继承边框表格的识别均有回归用例守护 |
| 回归基线 | `scripts/Invoke-RegressionBaseline.ps1` 一键跑三条链路（修复后复检 0 条、二次运行幂等、继承边框识别），留痕到 `artifacts/regression/`；`dotnet test` 当前实测 595 条全过 |
| 独立网页前端 | 当前产品不提供；用户工作界面是 Word 内的任务窗格 |
| 南开大学正式格式规则的完整覆盖 | 尚未完成；当前演示使用合成规则 |

CLI 标注演示会生成带批注和高亮的新 Word 文档；VSTO 任务窗格则在 Word 中提供诊断列表、定位和安全修复操作。两者是不同的演示入口。CLI 运行记录见[Demo 运行与 Word 验证记录](docs/demo/README.md)；加载项安装和使用见 [VSTO 插件说明](src/ThesisAddIn.VSTO/README.md)。独立网页前端不在当前产品范围内。

无标题样式、全为普通段落的 DOCX 处理步骤、可选 AI、大纲确认和完整合成稿见[无样式论文工作流](docs/unstyled-thesis-workflow.md)，开源选型与许可证见[调研记录](docs/research/unstyled-thesis-open-source-2026-10-26.md)。

## 架构概览

- **ThesisEngine.Core**：提供文档读取、规则检查、诊断和标注输出等核心能力。
- **ThesisEngine.Cli**：在本地运行检查与演示，输出中文诊断。
- **Word 桌面集成**：VSTO 插件提供用户界面和 Word 交互；按插件说明在 Windows 桌面版 Microsoft Word 中安装使用。
- **规则与契约**：规则描述检查要求，契约定义模块间的数据结构；正式规则须核验来源、版本和适用范围后才能采用。

## 环境要求

- .NET 10 SDK
- Microsoft Word 桌面版（仅在 Word 中查看标注演示时需要）

以下命令在仓库根目录运行。

## 构建与测试

```powershell
dotnet restore NkuThesisTool.sln
dotnet build NkuThesisTool.sln --configuration Release --no-restore
dotnet test NkuThesisTool.sln --configuration Release --no-build
```

改完格式化逻辑后，用一条命令跑回归基线（三条链路，输出留痕到 `artifacts/regression/`）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Invoke-RegressionBaseline.ps1
```

## 聊天模型配置（chat 子命令）

`chat` 子命令回答论文**格式**问题，支持离线规则与多个免费模型通道。格式范围由代码强制约束，API Key 只从环境变量读取（禁止硬编码）。

```powershell
# 离线模式，无需任何 Key
dotnet run --project src/ThesisEngine.Cli -- chat --local 一级标题用几号字

# 模型通道：先在环境变量配置至少一个免费 Key，例如
$env:DEEPSEEK_API_KEY = "sk-xxxx"
dotnet run --project src/ThesisEngine.Cli -- chat 一级标题用几号字

# 查看帮助与配置示例
dotnet run --project src/ThesisEngine.Cli -- chat --help
```

默认配置 `src/ThesisEngine.Cli/chat.models.json` 已列出 3 个免费通道（SiliconFlow / DeepSeek / Kimi）加 1 个 OpenAI 兼容接口模板；主通道失败会自动降级到另外 2 个免费通道。完整说明与排查见 [聊天模型配置](docs/chat-models.md)。

## 运行演示

合成的正确样例在段落级演示规则下预期得到 **0 条诊断**：

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/clean_sample.docx
```

`analyze` 默认 `--scope all`，段落级报告后还会附一份文档级版式报告。该极简样例沿用 Word 默认页面设置，文档级会报 8 处；只想看段落级加 `--scope paragraph`。注意这份样例的 0 条是和「已检查 0 个正文文本片段」一起出现的，属于「压根没查」，不是合规结论。

想看检查器**真正走完两级链路**且结果都干净，用这一份：

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/clean_checked_sample.docx
```

它会同时给出「已检查 3 个正文文本片段，诊断 0」和文档级「完整判定 2 / 2、已检查的文档级项目均符合」。样本的生成方式、Word 16 实机读数见[合成样例说明](tests/fixtures/synthetic/demo/README.md)。

含有已知格式差异的合成样例预期得到 **3 条诊断**：

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx
```

要生成可在 Word 中查看的标注副本，先创建输出目录，再运行：

```powershell
New-Item -ItemType Directory -Force artifacts | Out-Null
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx --annotated-output artifacts/known_issues.annotated.docx
```

生成后用 Word 打开 <code>artifacts/known_issues.annotated.docx</code>，可查看三处黄色高亮及对应中文批注。输入样例不会被改写，已有的输出文件不会被覆盖。演示样例检查东亚字体、西文字体和字号三项；完整预期结果和复测说明见[Demo 文档](docs/demo/README.md)。

## 参考文献一致性检查（只查不改）

`analyze --citations` 单独检查正文引用与文末参考文献是否对得上，**只报问题、不改写文档**。当前口径是顺序编码制（`[1]`、上标 `[1]`），依据及检查项见 [#152](https://github.com/lzmp666/nku-thesis-tool/issues/152)。

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --citations <论文.docx>
```

报告给出引用标注数、参考文献条目数、问题数及逐条定位（`word/document.xml#paragraph[n]`）。检查五类：正文引用编号连续（跳号/重号）、有引无条与有条无引、标注位置（标点之前）、条目字段完整性（编号/作者/年卷期页）、同一文献复用同一编号。上标引用按 `w:vertAlign="superscript"` 识别；正文里作参数说明或数组下标的 `[1]` 不误报。

注意：只检查规则明确要求的项，规则未覆盖的不作判断；Zotero / EndNote 插入的域只统计不解析；页下注体例、著者-出版年制不在本模式范围。合规样本预期 0 条问题，输出会明确「问题数量：0」。

## 摘要格式检查（只查不改）

`analyze-abstract` 检查中英文摘要标题与正文的字体、字号、加粗、对齐、行距、段前段后，以及关键词标签是否加粗，**只报问题、不改写文档**。依据 `rules/nku-graduate/general.yaml` 的 R4/R5（4.5 表4.4）：中文摘要标题黑体 18 磅加粗居中、单倍行距、段前 24 / 段后 18 磅，正文宋体 12 磅、固定值行距 20 磅、段前后 0 磅；英文摘要标题 Arial 18 磅，正文 Times New Roman 12 磅，其余同中文。

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-abstract <论文.docx>
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-abstract <论文.docx> --json
```

报告列出识别到的摘要段落、逐条定位（第 N 段）、期望值与实际值，以及「未检查 / 未判定」项。只按「摘要」「Abstract」标题文字定位摘要区，定位不到就写「未识别」，不假装合规；样式链读不到的维度列进「未判定」，不拿默认值凑问题数。该规则状态为 **draft**（来源 S1 2025 版条文，待 2026 版全文核验），报告会提示需人工确认。样本与问题码见 [摘要样本说明](tests/fixtures/synthetic/abstract/README.md)。

## 尾注部件格式检查（只查不改）

`analyze-back-matter` 检查参考文献、附录、符号注释表与个人简历及成果、勘误页与致谢的**正文**格式，**只报问题、不改写文档**。依据 `rules/nku-graduate/general.yaml` 的 R14–R17（4.8 表4.7）：参考文献宋体 / Times New Roman 10.5 磅、固定值行距 16 磅、段前后 0 磅；附录宋体 / Times New Roman 12 磅、两端对齐、首行缩进 2 汉字符、固定值行距 20 磅、段前后 0 磅；符号表与简历 10.5 磅 / 固定 16 磅；致谢仿宋 12 磅 / 固定 16 磅。

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-back-matter <论文.docx>
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-back-matter <论文.docx> --json
```

问题码按部件连续编号：参考文献 B1–B4、附录 B5–B10、符号表与简历 B11–B14、致谢 B15–B17。报告同时列出识别到的部件与正文段落、期望值与实际值、未检查项。口径边界：各部件标题要求「同各章标题（R7）」由标题检查负责，本命令不判；附录中含数学表达式的段落按条文豁免行距，其余维度照查；同部件子标题（附录 A/B）不算正文。识别不到标题的部件列入「未识别」，不假装合规；样式链读不到的维度列入「未判定」。规则状态为 **draft**（来源 S1 2025 版条文），样本与问题码见 [尾注样本说明](tests/fixtures/synthetic/back-matter/README.md)。

本命令与 `analyze --citations` 的参考文献**引用一致性**（#152）是两个不同维度：这里只管字体字号行距等排版值，不核对正文引用与文末条目是否一一对应。

## 正文段落格式检查（R11 补齐，只查不改）

`analyze-body-text` 补齐 R11 中 #43 尚未覆盖的四个排版维度：两端对齐、首行缩进 2 汉字符、固定值行距 20 磅、段前后 0 磅，**只报问题、不改写文档**。依据 `general.yaml` 4.7 表4.6（来源 S2026 2026 版，#43 已核对）。字体与字号（宋体 / Times New Roman 12 磅）仍由 `analyze-styles` 的 #43 规则负责，本命令**不重复判定**。

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-body-text <论文.docx>
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-body-text <论文.docx> --json
```

正文区间取第一个章标题之后、参考文献/附录/致谢等结束标题之前；边界识别不到就如实列未检查，不拿封面/摘要冒充正文。标题、图表题注、表格内段落不是正文，跳过。问题码 T1–T4（对齐/缩进/行距/段距）。含数学表达式的段落按条文豁免行距，其余维度照查。规则状态 **draft**，样本见 [正文样本说明](tests/fixtures/synthetic/body-text/README.md)。

## 标题层级格式检查（R7–R10，只查不改）

`analyze-headings` 检查章/一级/二级/三级标题的字体、字号、加粗、单倍行距、段前段后，**只报问题、不改写文档**。依据 `general.yaml` 4.7 表4.6（来源 S4 2025 版 / S1 2026 版，规则状态 draft）：章黑体 16 磅加粗、段前 24 / 段后 18 磅；一级 14 磅加粗、段前 24 / 段后 6；二级 13 磅、段前 12 / 段后 6；三级 12 磅、段前 12 / 段后 6。标题级别只按 `w:outlineLvl` 或 `HeadingN` /「标题 N」样式识别，不靠字体猜测；识别不到就如实列未检查。

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-headings <论文.docx>
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-headings <论文.docx> --json
```

问题码 H1–H22 连续（每级 CodeStart + 维度偏移，避免拼出 H111）。对齐依编号模式、序号间空一汉字宽度语义未定义、二级/三级原文未写加粗——三类按条文冲突列**未判定**，不伪造硬规则。样本见 [标题样本说明](tests/fixtures/synthetic/headings/README.md)。

## 目录格式检查（R6，只查不改）

`analyze-toc` 检查目录标题与章/一级/二级目录条目的字体、字号、单倍行距、段前后、左缩进，**只报问题、不改写文档**。依据 `general.yaml` 4.6 表4.5（来源 S1 2025 版，规则状态 draft）：目录标题黑体 16 磅加粗居中、段前 24 / 段后 18 磅；章/一级/二级条目宋体 14/12/10.5 磅、段前 6 / 段后 0 磅、左缩进 0/1/2 汉字。

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-toc <论文.docx>
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-toc <论文.docx> --json
```

目录标题用文本「目录 / Contents」定位，条目按 `TOC1` / `TOC2` / `TOC3` 样式分级，不靠缩进硬猜。检测对象是**已展开的目录段落**：只有 TOC 域指令、域结果未展开时列**未判定**（检测对象 T6 / RFC-7 未决，需 Word 更新域后复核），不假装检查过。条目的页码右对齐走制表位、依赖渲染，列未判定。问题码 C1–C25 连续。样本见 [目录样本说明](tests/fixtures/synthetic/toc/README.md)。

## 统一合规报告（只查不改）

`report` 一次调用跑完全部「只查不改」分析器（页面与章节结构、样式使用、图表排版、公式排版、脚注与注释体例、参考文献引用一致性、中英文摘要格式、尾注部件格式、正文段落格式、标题层级格式、目录格式），汇成一份报告，**只报问题、不改写文档**。适合交付前一次性总览，不必逐条命令分别运行。

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- report <论文.docx>
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- report <论文.docx> --json
```

报告分三部分：**规范问题**、**排版惯例**（不计入规范符合性）、**未检查 / 未判定**——三者严格分列，判不出的不冒充合规。开头列出十一个检查器各自的执行情况；如果某个分析器因文档缺部件而失败，会明确写出错误信息，其余检查照常产出，失败的维度不据此判定合规。所有规则状态均为 **draft**，报告为静态 XML 检查，不替代 Word 实机渲染核验。

## 在 Word 中打开插件

完成本机签名和当前用户注册后，重新启动 Microsoft Word，打开一份合成 `.docx`。
顶部的 **论文格式检查** 功能区提供分析、单项/批量修复、问题与位置导航，以及「问题检索」列表开关。

点击「分析当前文档」会直接分析并更新当前打开的文档，不另开文档。右侧「问题检索」仅保留搜索框和按统一目标格式合并的问题列表；详细说明显示在对应的 Word 批注中。确认修复后，改动直接留在当前 Word 文档，可撤销；作者完成后自行保存。

## 当前范围与限制

- 只检查演示规则覆盖到的 Word 正文普通段落和文本片段；位置是 OOXML 部件及段落/文本片段索引，不是页码。
- 当前合成规则包括东亚字体、西文字体和字号。即使报告为零问题，也只表示样例中这三项规则没有发现差异。
- 表格、页眉页脚、脚注尾注、超链接、修订内容、内容控件和无法解析的主题字体等尚未纳入演示覆盖范围。
- 同一目标格式的问题及批注合并为一个诊断项，并保留每处位置。确认修复后直接修改当前打开的文档，不会逐项另建或重开结果文档。东亚字体、已解析的 ASCII/HAnsi 西文字体和字号都由修复引擎支持；插件先展示修改预览并要求作者确认。未知格式、无法安全定位的范围仍停止修复。
- 当前插件仍按本地开发方式注册，尚未提供面向成员的正式安装包或部署流程。
- 当前规则仅为合成演示规则，不承诺正式规范合规。正式规则需先对照学校发布的规范核验。

## 隐私与样例

仓库中的演示文档为人工构造的合成样例，不含真实论文内容。请勿将未经授权的论文、个人信息、访问令牌、密码或其他敏感材料提交到仓库。提交前请检查改动、附件和 Git 历史中是否包含这类内容。

## 参与贡献

项目采用 **Issue → Branch → PR → Review → Merge** 流程：先阅读 Issue 的目录边界和验收标准，再领取任务、创建短期分支、完成自测并提交 PR；不要直接向 <code>main</code> 推送。详见[贡献指南](docs/contribution.md)、[规则指南](docs/rule-guide.md)和[QA 流程](docs/qa-sop.md)。实际分支保护状态以仓库设置为准。

## 许可证

本项目采用 [MIT License](LICENSE)。

