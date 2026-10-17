# 论文格式检查 Demo 运行说明

> **仅供草稿/演示。** 本演示使用合成规则和合成样例，不是南开大学现行规范检查，也不构成合规结论。请勿使用真实论文或个人信息运行演示。

## 环境要求

- .NET 10 SDK
- 已检出本仓库代码

以下命令请在仓库根目录运行。

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

## 输出标注副本

在命令末尾添加 `--annotated-output <标注副本.docx>`，会将诊断写入新的 Word 副本，包含批注和黄色高亮；输入文件保持不变，已有目标不会被覆盖。输出目录必须事先创建。

```sh
mkdir -p artifacts
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx --annotated-output artifacts/known_issues.annotated.docx
```

命令报告会额外列出标注副本路径。无问题样例仍会生成副本，但不会添加批注或高亮。

## 当前检查范围

读取器检查 Word 正文中的普通段落和文本片段。它会解析直接格式、段落和字符样式的 `basedOn` 继承链，以及文档默认设置中的东亚字体、ASCII/HAnsi 字体和字号。位置使用 OOXML 文档部件、段落序号和文本片段序号表示，**不是页码**。

这是一个范围很小的演示。它不代表符合任何现行高校规范，也不会自动修复文档。对于表格、页眉/页脚、脚注/尾注、超链接、修订内容、内容控件和无法解析的主题字体等未覆盖内容，报告会给出提示。即使报告没有发现问题，也只表示这三项演示规则在已检查文本片段中没有发现不匹配。

## 常见问题

- 如果提示找不到样例文件，请先切换到仓库根目录再运行命令。
- 如果构建输出中没有 CLI 项目，请先执行上面的 Release 构建命令，再使用 `--no-build` 运行演示。
- 如果还没有安装 SDK，请安装 .NET 10 SDK 后重新执行还原命令。
- 如果检查结果与表格中的三项不一致，请按[样例说明](../../tests/fixtures/synthetic/demo/README.md)重新生成合成样例，再运行测试。

## 演示图片

[查看中文 CLI 输出示例](demo-output-zh.svg)。这是根据端到端测试验证的演示结果制作的静态呈现，不是实时屏幕截图，部分中间字段有所省略。完整报告请运行上面的命令。图片只包含合成样例路径和演示结果。

## Word 打开验证记录（Issue #46）

验证日期：2026-10-15　验证人：@piggy515

### 验证环境

| 项目 | 版本 / 说明 |
| --- | --- |
| 操作系统 | Windows 11 25H2，内部版本 26200.9457（注册表产品名仍显示 Windows 10，为本机已知显示现象） |
| Microsoft Word | 16.0.20430.20146（桌面版；本次账户页显示版本 2609、内部版本 20430.20146，即点即用） |
| .NET SDK | 10.0.401（原 Release 构建与完整测试）；10.0.203（本次 CLI 与 Word 界面复测） |
| 授权状态 | 本机 Windows / Office 未激活，Word 标题栏显示“未经授权产品”并提示“重启以获得 Office”。该提示来自 Office 授权状态，与本文档无关，未影响打开与解析。 |

### 运行命令与输出路径

在仓库根目录执行：

```sh
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx --annotated-output artifacts/known_issues.annotated.docx
```

输出路径：`artifacts/known_issues.annotated.docx`。`artifacts/` 与 `*.docx` 均被 `.gitignore` 忽略，标注副本不随仓库提交，可用上述命令随时重新生成。运行前需先创建 `artifacts` 目录。

### 验收结果

1. **可打开**：标注副本在 Microsoft Word 中正常打开，未出现“发现无法读取的内容 / 是否修复”一类提示，未触发修复流程。
2. **定位与高亮**：3 处诊断均落在对应段落上，并以黄色高亮显示。

| 序号 | 规则编号 | 位置 | 期望值 | 实际值 |
| --- | --- | --- | --- | --- |
| 1 | `demo/east-asian-font` | `word/document.xml#paragraph[0]/run[0]` | 宋体 | 黑体 |
| 2 | `demo/western-font` | `word/document.xml#paragraph[1]/run[0]` | Times New Roman | Arial |
| 3 | `demo/font-size` | `word/document.xml#paragraph[2]/run[0]` | 12pt | 10.5pt |

3. **中文批注**：文档内含 3 条批注，作者均为 `NKU Thesis Tool`。按 Issue #63 的要求，每条批注只保留四项信息——要求参考来源、规则、期望值、实际值，不再重复”检测提示”或”修改建议”等同义说明。逐条如下：

   - `要求参考来源：2026版第4.7节表4.6；规则：demo/east-asian-font；期望值：宋体；实际值：黑体。`
   - `要求参考来源：2026版第4.7节表4.6；规则：demo/western-font；期望值：Times New Roman；实际值：Arial。`
   - `要求参考来源：2026版第4.7节表4.6；规则：demo/font-size；期望值：12pt；实际值：10.5pt。`

   每条批注通过 `commentRangeStart`/`commentRangeEnd`/`commentReference` 与对应黄色高亮文本范围锚定，三者使用同一批注 ID。
4. **正确样例**：`clean_sample.docx` 生成的副本不含高亮与批注，与”0 条诊断”一致。
5. **输入不变**：运行前后逐一比对输入文件字节，内容未发生变化。
6. **构建与测试**：Release 构建 0 警告 0 错误；解决方案全部测试通过（63 项）。

### 说明与限制

- **Word 界面复核（2026-10-15，@lzmp666）**：本次使用 .NET SDK 10.0.203 按上方命令重新生成 `artifacts/known_issues.annotated.docx` 并在桌面版 Word 中打开，未出现“发现无法读取的内容 / 是否修复”提示。Word 窗口较窄、右侧扩展窗格占用空间时，批注标记区未进入可视范围；关闭右侧扩展窗格并最大化窗口后，右侧可见 3 条由 `NKU Thesis Tool` 发布的中文批注，与 3 处黄色高亮一一对应。Word 同时显示“未经授权产品”及“兼容模式”，但本次实测中批注仍正常显示。此前“未能直接确认批注界面可见”的描述已由本次界面复测更正。
- **正确样例复测**：另以 `clean_sample.docx` 运行 CLI 并输出 `artifacts/clean.annotated.docx`，得到 0 条诊断；该副本没有 `word/comments.xml`，正文没有高亮元素。两份输入样例均与仓库原始样本的 SHA-256 一致，输入未被改写。
- **批注精简与交互限制（Issue #63，2026-10-16）**：批注正文已精简为四项（要求参考来源、规则、期望值、实际值），不再附”检测提示/修改建议”等重复说明。在 Word 16.0.20430 中实测：点击右侧批注卡片会高亮对应原文；但点击正文黄色高亮仅移动插入点，不会反向激活对应批注卡片。正文到批注的双向联动受 Word/VSTO 原生行为限制，后续在 Issue #51 中以中文侧边栏方式实现，本阶段不把”悬停高亮定位”描述为已验证能力。
- Word 以”兼容模式”打开该标注副本。本次不影响验收结论；是否需要为生成的文档补充兼容性设置，建议作为后续项单独评估。
- 本记录只覆盖演示范围内的三项规则（中文字体、西文字体、字号）。表格、页眉页脚、脚注尾注、超链接、修订、内容控件及无法解析的主题字体等仍未覆盖，详见上文“当前检查范围”。未覆盖项不因本次没有诊断或诊断已标注而被视为合规。
