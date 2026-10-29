# 南开大学研究生学位论文格式检查与整改工具

面向论文排版检查与辅助整改的开源原型。产品交互界面运行在桌面版 Word 的 VSTO 任务窗格中；.NET 核心引擎和 CLI 用于分析、自动化测试及演示样例。当前只覆盖少量合成格式规则，不是南开大学正式规范的完整实现，也不代表论文合规结论。

> 当前仓库已公开。演示请使用仓库中的合成样例；不要提交真实论文、个人信息或未经授权的文档。

## 当前进度

| 能力 | 状态 |
| --- | --- |
| .NET 10 核心引擎与 CLI | 可运行 |
| 无样式论文结构与角色排版 | Word 1.5.4 预览：大纲确认、明确勾选范围、原位整篇排版、写后回读与整篇撤销；核心/CLI 保留可选 Ollama 接口 |
| 中文诊断报告 | 可运行 |
| 将诊断写入新的 Word 文档副本（批注和黄色高亮） | 可运行；已在 Microsoft Word 桌面版打开验证 |
| Word 插件前端（中文任务窗格、原文定位、锚点联动、单项/批量修复预览） | 已合入 main（#65、#69）；已有合成文档实测记录，完整阶段验收仍由 #51、#63、#52 跟踪 |
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

## 运行演示

合成的正确样例预期得到 **0 条诊断**：

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/clean_sample.docx
```

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

