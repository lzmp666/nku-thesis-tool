# AGENTS.md

在这个仓库里干活，先读完这一页。写它是为了让 AI 一上来就知道边界——能改哪里、不能改哪里、跑什么命令、交付什么。人也能照着走。
更细的上手步骤见 [docs/ai-onboarding.md](docs/ai-onboarding.md)，规则怎么落成 YAML 见 [docs/rule-guide.md](docs/rule-guide.md)。

## 1. 项目概览

南开大学研究生学位论文的格式检查与整改工具。引擎读 OOXML，出中文诊断，能改的地方只改格式。

- `ThesisEngine.Core`：读 DOCX、按规则检查、生成补丁、写诊断、撤销。
- `ThesisEngine.Cli`：命令行入口，本地跑检查和演示。
- `ThesisAddIn.VSTO`：Word 任务窗格插件，用户实际操作的界面在这一层。
- `rules/` 与 `contracts/`：规则数据、模块间数据结构。

**检查结论只对已加载的规则有效。** 「0 条问题」表示这些规则没发现差异，不表示论文合规。报告里必须能区分「查了没发现」和「压根没查」——没查的项要列出来，不能省略。

## 2. 第一原则

**格式全部走 Word 样式，不逐段刷直接格式，不改文字内容。**

- 改样式定义，全文跟着变；改单段直接格式，章节一多就查不清是谁动的。
- 正文文字、标点、段落顺序、表格内容、公式对象里的 `m:t`，一个字符都不许改。这是排版工具，不是改稿工具。
- 要「看起来对齐」时，先想样式、制表位、段落属性，不要敲空格凑。
- 期望值只能来自校方条文，不能来自手感。进 `rules/` 之前核验出处、版本、适用范围。
- 判不准就报「未检查」并写明原因，不能写「符合」。
- 修完要复诊：重新跑一遍检查，确认问题消失且没冒出新问题。`auto` 级修复必须验证语义幂等。

## 3. 构建与测试

需要 .NET 10 SDK（`global.json` 锁 10.0.100，允许 `latestFeature` 上浮）。Word 桌面版只在看标注副本或装插件时需要。

```powershell
dotnet restore NkuThesisTool.sln
dotnet build NkuThesisTool.sln --configuration Release --no-restore
dotnet test NkuThesisTool.sln --configuration Release --no-build
```

`Directory.Build.props` 开了 `TreatWarningsAsErrors`，冒出一个警告就是构建失败，别想着先放着。
基线（2026-11-11，main）：build 0 警告 0 错误，test 459 条全过。这个数会随 main 涨，别写死进代码或断言总数。

## 4. 目录边界

| 路径 | 干什么 | 改之前 |
| --- | --- | --- |
| `src/ThesisEngine.Core/` | 引擎本体，不依赖 UI | CODEOWNERS 要求负责人审查 |
| `src/ThesisEngine.Cli/` | 命令行，一个命令一个 `*Command.cs` | 新命令在 `Program.cs` 加一行分发 |
| `src/ThesisAddIn.VSTO/` | Word 插件，Windows + .NET Framework | 先读 `src/ThesisAddIn.VSTO/README.md` |
| `tests/ThesisEngine.Tests/` | xUnit 测试 | 只加不删，别改已有断言迁就实现 |
| `tests/fixtures/synthetic/<主题>/` | 合成样本 + `generate.py` + `manifest.json` | 改样本就重跑 `generate.py`，docx 和 manifest 一起提交 |
| `rules/nku-graduate/` | 规则 YAML，目前是 draft | 每条带出处、版本、适用范围 |
| `contracts/` | 数据结构、`rule-schema.yaml`、RFC | 先提 RFC，等负责人确认 |
| `docs/` | 流程、调研、验收记录 | 可以加，不要重写别人的正文 |

不许动：`.github/`（CI、模板、CODEOWNERS）、`Directory.Build.props`、`global.json`、`NkuThesisTool.sln`、`NuGet.Config`。这些是工程配置，需要改就开 issue 说清理由，由负责人改。

## 5. 命令

以下都在仓库根目录、且已经 build 过之后跑。写入型命令的输出目录要先建：`New-Item -ItemType Directory -Force artifacts | Out-Null`。

```powershell
# 演示诊断：极简正确样本 0 条（但已检查 0 个片段，属「压根没查」）、问题样本 3 条；默认 --scope all 还会附文档级报告（极简样本会报 8 处）
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx
# 真正走完两级链路的干净样本：段落级已检查 3 个片段且 0 条，文档级完整判定 2/2 且均符合
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/clean_checked_sample.docx
# 只跑文档级：--scope all|paragraph|document，默认 all 把文档级报告附在段落级后面
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx --scope document
# 写一份带批注和黄色的标注副本（输出目录要先存在，已有文件不覆盖）
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx --annotated-output artifacts/known.annotated.docx
# 样式体系：哪些段落没套样式、套了样式又被直接格式覆盖
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-styles tests/fixtures/synthetic/demo/clean_sample.docx
# 图表：图片/表格数量与排版问题
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-figures tests/fixtures/synthetic/full-thesis/synthetic_object_layout_errors.docx
# 公式：居中式号字体字号（F1–F5）
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-formula tests/fixtures/synthetic/formula/formula_word_native_correct.docx
# 文档级：纸张页边距页眉页脚页码目录与篇眉页码字体（D1–D12）
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze-document tests/fixtures/synthetic/document/document_baseline_correct.docx
# 套用南开格式：只改样式定义，不给输出名就加「_nku格式」后缀
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- apply-nku-format tests/fixtures/synthetic/demo/clean_sample.docx artifacts/run_nku.docx
# 图表格式写入：三线表、图片居中、跨页保护
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- apply-figure-format tests/fixtures/synthetic/full-thesis/synthetic_object_layout_errors.docx artifacts/run_fig.docx
# 公式格式写入：居中、式号右对齐、Cambria Math、段距归零
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- apply-formula-format tests/fixtures/synthetic/formula/formula_number_not_right_aligned.docx artifacts/run_formula.docx
# 文档级格式写入：只动 sectPr 的纸张与页边距，改完自动复诊
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- apply-document-format tests/fixtures/synthetic/document/document_margin_wrong.docx artifacts/run_doc.docx
# 一步到位：备份 + 套样式 + 改图表
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- full-format tests/fixtures/synthetic/demo/clean_sample.docx artifacts/run_full.docx
# 撤销：按输出文件定位，先备份到系统临时目录的 nku-thesis-tool-backups
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- undo artifacts/run_nku.docx
# 规则问答：没配 DEEPSEEK_API_KEY 就走本地规则
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- chat "正文行距多少"
# 结构识别：出角色候选和大纲，--format-output 才写入
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- structure tests/fixtures/synthetic/full-thesis/synthetic_complete_thesis.docx --report artifacts/structure.json
```

写入型命令（`apply-*`、`full-format`、`structure --format-output`、`analyze --annotated-output`）一律另存新文件，不改输入。`apply-document-format` 只写 `sectPr` 下的 `pgSz` / `pgMar`，目录、页码、空白页只报告不改。默认输出名带中文后缀，要精确控制就显式给第二个参数。

改过 `tests/fixtures/synthetic/` 下的样本或生成器，按顺序跑这三条（前两条只用标准库，不需要 pip install）：

```bash
python3 tests/fixtures/synthetic/verify_fixtures.py                 # 默认档：零依赖
python3 tests/fixtures/synthetic/verify_fixtures.py --self-test     # 验证判据本身
python3 tests/fixtures/synthetic/verify_fixtures.py --with-rebuild  # 加严档：重建 12 份样本
```

第一条核对逐部件重建（4 个零依赖生成器 / 39 份样本）、结构断言（format-samples 10 份）与 manifest ↔ 生成器 ↔ 磁盘的清单一致性（5 个目录 / 51 份 docx）；第二条用 12 种情形验证这些判据既不误报也不漏报。

`--with-rebuild` 把 `format-samples/` 的 10 份和 `full-thesis/` 的 2 份也做一遍逐部件重建，合计 12 份。它需要 `python-docx`（full-thesis 另需 `Pillow`、`lxml`），版本见两个目录各自的 `requirements.txt`。缺依赖时**整组**跳过并打印原因，退出码仍是 0——「跳过」不等于「通过」，别把跳过当验收。`full-thesis/` 的图片文字取自 Windows 字体 `msyh.ttc`，换平台或换字体后重建结果的图片字节会变，属预期。

两份 `full-thesis` 样本由 `python3 tests/fixtures/synthetic/full-thesis/generate.py --check` 核对，比的是逐部件而不是整文件哈希：python-docx 保存时 ZIP 时间戳取当前时刻，整文件字节每次不同。

## 6. 提交规范

Issue → 分支 → PR，一步不能跳。直接推 `main` 视为没做。

1. 在 issue 留言领取，等负责人设 assignee 再动手。一个人先领一个。
2. 分支前缀只用 `feature/`、`fix/`、`rules/`、`test/`、`docs/`。一个 issue 一个分支，不建 `develop`。
3. 只改 issue 里「允许修改目录」列出的路径。想越界，先在 issue 里说明，等确认。
4. 提交前跑完 build 和 test，把真实输出留着。
5. PR 正文按 `.github/pull_request_template.md` 填，`Closes #编号`。**必须贴**：build 的警告/错误数、test 的通过数、至少一条命令的实跑输出。写「本地测试通过」不算。
6. `contracts/` 和 `src/ThesisEngine.Core/` 的改动要负责人 Review（看 CODEOWNERS）。

不清楚的边界以 `.github/ISSUE_TEMPLATE/task.md` 的字段为准：背景、工作内容、允许修改目录、不允许修改目录、验收标准、前置依赖、领取。

## 7. 验收清单

提 PR 前逐条过：

- [ ] 只改了 issue 允许的目录，没碰 `.github/`、工程配置、`sln`
- [ ] `dotnet build` 0 警告 0 错误（警告也算失败）
- [ ] `dotnet test` 全过，新增行为有对应测试
- [ ] 新样本由生成脚本产出（最好能 `--check` 逐字节复现），`manifest.json` 同步更新，没直接提交手改的 docx
- [ ] 命令是我真跑过的，文档里的参数没有编的
- [ ] 规则的期望值有出处（条、表、版本），draft 没冒充正式规则
- [ ] 报告区分了「查了没问题」和「没查」，没查的项有名字
- [ ] 修完复诊过，二次检查为 0，且没引入新问题
- [ ] 没有真实论文、个人信息、密钥、`bin/`、`obj/`

## 8. 已用到的开源实现与许可证

本项目 MIT。依赖清单：

| 依赖 | 版本 | 许可证 | 用在哪 |
| --- | --- | --- | --- |
| DocumentFormat.OpenXml | 3.5.1 | MIT | 仅测试工程，造样本、核对结构。引擎本身不用它，直接读 XML |
| xunit | 2.9.3 | Apache-2.0 | 测试框架 |
| xunit.runner.visualstudio | 3.1.1 | Apache-2.0 | 测试运行器 |
| Microsoft.NET.Test.Sdk | 17.14.1 | MIT | 测试宿主 |
| System.Text.Json | 8.0.5 | MIT | 仅 VSTO 协议工程（.NET Framework 没有内置） |
| python-docx | 1.2.0 | MIT | 仅 `tests/fixtures/synthetic/` 下 `format-samples/generate.py` 与 `full-thesis/` 的重建路径；只读校验不用它 |
| Pillow | 13.0.0 | MIT-CMU | 仅 `full-thesis/generate_full_thesis.py` 画合成图表；`--with-rebuild` 用它重建 |
| lxml | 6.1.3 | BSD-3-Clause | 仅 `full-thesis/generate_object_layout_errors.py` 改派生样本的 XML |

后三个都是可选依赖，不是引擎运行前提：`generate.py --check-only`、`full-thesis/generate.py --check` 与 `tests/fixtures/synthetic/verify_fixtures.py` 的默认档都只用 Python 标准库，干净检出后不必 `pip install` 就能核对已提交样本。`verify_fixtures.py --with-rebuild` 才需要它们，缺了就整组跳过并说明。

`AGENTS.md` 的七段骨架参考 `AfishInLake/WordFormat`（Apache-2.0）的组织方式，只借结构，没有拷贝代码。**后续谁想重写一遍，先看这张表**：能直接用的别再自己造，抄了别人的要在这里补一行。

## 9. 相关文档

- [docs/contribution.md](docs/contribution.md)：从领取到合并的完整流程
- [docs/architecture.md](docs/architecture.md)：数据流与写入边界
- [docs/qa-sop.md](docs/qa-sop.md)：自测怎么做
- [docs/demo/README.md](docs/demo/README.md)：演示的预期输出
- [docs/rule-notes/page-setup.md](docs/rule-notes/page-setup.md)：文档级检查的判断口径
- [docs/research/word-112-acceptance.md](docs/research/word-112-acceptance.md)：Word 实机核验读数与截图
