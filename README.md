# 南开大学研究生学位论文格式检查与整改工具

面向论文排版检查与辅助整改的开源原型。项目目前提供 .NET 核心引擎和命令行演示：可检查少量合成格式规则、生成中文诊断，并把问题写成 Word 批注和黄色高亮。当前演示不是南开大学正式规范的完整实现，也不代表论文合规结论。

> 当前仓库已公开。演示请使用仓库中的合成样例；不要提交真实论文、个人信息或未经授权的文档。

## 当前进度

| 能力 | 状态 |
| --- | --- |
| .NET 10 核心引擎与 CLI | 可运行 |
| 中文诊断报告 | 可运行 |
| 将诊断写入新的 Word 文档副本（批注和黄色高亮） | 可运行；已在 Microsoft Word 桌面版打开验证 |
| Word 插件、任务窗格及单条或批量修复 | 尚未完成 |
| 南开大学正式格式规则的完整覆盖 | 尚未完成；当前演示使用合成规则 |

当前 Word 标注演示由 CLI 生成 Word 文档，再在 Word 桌面版中查看；它不是 Word 加载项界面。运行记录、复测说明和边界见[Demo 运行与 Word 验证记录](docs/demo/README.md)。Word 集成进展见 [Issue #50](https://github.com/lzmp666/nku-thesis-tool/issues/50)。

## 架构概览

- **ThesisEngine.Core**：提供文档读取、规则检查、诊断和标注输出等核心能力。
- **ThesisEngine.Cli**：在本地运行检查与演示，输出中文诊断。
- **Word 桌面集成**：后续以 VSTO 插件接入 Word；当前主线尚未提供可用的 Word 加载项。
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

## 当前范围与限制

- 只检查演示规则覆盖到的 Word 正文普通段落和文本片段；位置是 OOXML 部件及段落/文本片段索引，不是页码。
- 当前合成规则包括东亚字体、西文字体和字号。即使报告为零问题，也只表示样例中这三项规则没有发现差异。
- 表格、页眉页脚、脚注尾注、超链接、修订内容、内容控件和无法解析的主题字体等尚未纳入演示覆盖范围。
- 当前不提供单条修复或一键修复功能，也不承诺正式规范合规。正式规则需先对照学校发布的规范核验。
- Word 批注与高亮是 CLI 写入新文档副本的演示结果，不能据此推断 Word 插件任务窗格已经完成。

## 隐私与样例

仓库中的演示文档为人工构造的合成样例，不含真实论文内容。请勿将未经授权的论文、个人信息、访问令牌、密码或其他敏感材料提交到仓库。提交前请检查改动、附件和 Git 历史中是否包含这类内容。

## 参与贡献

项目采用 **Issue → Branch → PR → Review → Merge** 流程：先阅读 Issue 的目录边界和验收标准，再领取任务、创建短期分支、完成自测并提交 PR；不要直接向 <code>main</code> 推送。详见[贡献指南](docs/contribution.md)、[规则指南](docs/rule-guide.md)和[QA 流程](docs/qa-sop.md)。实际分支保护状态以仓库设置为准。

## 许可证

本项目采用 [MIT License](LICENSE)。
