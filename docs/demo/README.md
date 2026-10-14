# 论文格式检查 Demo 运行说明

> **仅供草稿/演示。** 本演示使用合成规则和合成样例，不是南开大学现行规范检查，也不构成合规结论。请勿使用真实论文或个人信息运行演示。

## 环境要求

- .NET 10 SDK
- 已检出本仓库代码

以下命令请在仓库根目录运行。CLI 只读取输入文件，不会修改原文档。

## 构建和测试

```sh
dotnet restore NkuThesisTool.sln
dotnet build NkuThesisTool.sln --configuration Release --no-restore
dotnet test NkuThesisTool.sln --configuration Release --no-build
```

## 运行演示

正确的合成样例应输出 **0 条问题**：

```sh
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/clean_sample.docx
```

已知问题样例应输出 **恰好 3 条诊断**：

```sh
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx
```

报告会标明“草稿/演示”模式，并显示规则版本、检查覆盖范围、未覆盖范围和免责声明。预期的三项问题如下：

| 规则编号 | 位置 | 期望值 | 实际值 |
| --- | --- | --- | --- |
| `demo/east-asian-font` | `word/document.xml#paragraph[0]/run[0]` | 宋体 | 黑体 |
| `demo/western-font` | `word/document.xml#paragraph[1]/run[0]` | Times New Roman | Arial |
| `demo/font-size` | `word/document.xml#paragraph[2]/run[0]` | 12pt | 10.5pt |

期望值和样例内容记录在[合成样例清单](../../tests/fixtures/synthetic/demo/manifest.json)中。所有样例文字均为人工构造，不含真实论文或个人信息。

## 当前检查范围

读取器检查 Word 正文中的普通段落和文本片段。它会解析直接格式、段落和字符样式的 `basedOn` 继承链，以及文档默认设置中的东亚字体、ASCII/HAnsi 字体和字号。位置使用 OOXML 文档部件、段落序号和文本片段序号表示，**不是页码**。

这是一个范围很小的演示。它不代表符合任何现行高校规范，也不会自动修复文档。对于表格、页眉/页脚、脚注/尾注、超链接、修订内容、内容控件和无法解析的主题字体等未覆盖内容，报告会给出提示。即使报告没有发现问题，也只表示这三项演示规则在已检查文本片段中没有发现不匹配。

## 常见问题

- 如果提示找不到样例文件，请先切换到仓库根目录再运行命令。
- 如果构建输出中没有 CLI 项目，请先执行上面的 Release 构建命令，再使用 `--no-build` 运行演示。
- 如果还没有安装 SDK，请安装 .NET 10 SDK 后重新执行还原命令。
- 如果检查结果与表格中的三项不一致，请按[样例说明](../../tests/fixtures/synthetic/demo/README.md)重新生成合成样例，再运行测试。

## 演示图片

[查看合成样例的 CLI 输出图片](demo-output.png)。该图片是端到端测试验证过的输出节选的静态呈现，不是实时屏幕截图；中间报告内容有所省略。完整报告请运行上面的命令。图片只包含合成样例路径和演示结果。
