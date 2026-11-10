# 合成样本验证记录

## 2026-10-06 自审

- Python：使用工作运行时 Python/python-docx 1.2.0；生成后 --check-only 验证已提交的 10 个文件。
- 包与属性：10/10 ZIP 完整性、部件可读、文本、全部正文字体/字号/行距及所有节四边边距通过；继承对照的正文 run 无 rPr，覆盖错误仅第二正文段存在 eastAsia 属性。
- 重复生成：相同环境中另一个临时目录再生成，10/10 文件 SHA-256 一致。只用于生成可复现性验证，不用于格式修复幂等性判据。
- 负向校验：在临时副本中额外改坏字号错误文件的右边距，校验器正确拒绝，未修改提交的样本。
- CLI：现有 .NET 10 CLI 对 10 个样本均只读加载成功、退出码 0、无 stderr，原输入哈希不变。加载器不检测格式，不把该结果当格式检测通过。
- 项目回归：.NET SDK 10.0.203，Release restore/build/test 成功，0 warning/0 error，14/14 现有测试通过。

## Word 实测

独立隐藏的 Word 实例，通过 Documents.OpenNoRepairDialog 以 ReadOnly=true 打开，未指定强制修复、未保存；读取有效字体/字号/行距和节页面设置。Word 16.0，build 16.0.20430；Windows NT 10.0.26200.0。10/10 正常返回文档对象并取得预期有效属性。非截图 UI 人工复核，完整结果见 [word-validation.json](word-validation.json)。

- 普通正文：宋体/Times New Roman/12pt；各错误样本仅对应维度偏离。
- 行距：Word 返回基线规则 1、18pt（1.5 倍），行距错误正文规则 0、12pt（单倍）。
- 多节：第1节左 85.05pt；错误第2节左 70.85pt，其他边距不变。
- 在内存中把 Normal 中文字体改为黑体：继承对照第一/第二正文段均变为黑体；覆盖错误第一段变黑体、第二段仍楷体；关闭时不保存。

## 2026-10-06 尚未覆盖与审核状态

仅验证该 Word build 的打开和对象模型属性，没有验证其他 Windows/Mac Word、WPS、全部字体安装/字体替代或逐页视觉布局；这些范围未宣称通过。学校格式来源审核、正式引擎诊断/修复和第二人审核仍为 pending。此记录是自审证据，不代替非作者 approval。

## 2026-11-10 校验路径零依赖化

改动范围只有校验脚本与文档，**10 份 docx 字节未变**（`git status` 中无 `.docx` 变更）。

- 依赖：`python-docx==1.2.0` 的 import 从模块顶层下沉到 `_require_docx()`，只在重建路径触发。缺依赖时打印可执行的安装提示并指向 `--check-only`，不再抛裸 traceback。
- 新增只读 `verify(directory)`：用标准库 `zipfile` + `ElementTree` 复刻全部断言，返回问题描述列表，不写盘。`--check-only` 现在零依赖可用。
- 补齐漏检：原脚本只断言正文行距，标题行距从未被检查，而本目录 README 写着「标题仍 1.5 倍」。新增标题行距断言（`lineRule=auto`、`line=360`），与正文断言口径一致。
- 依赖许可证：python-docx 1.2.0 经 PyPI 元数据核实为 MIT（classifier `OSI Approved :: MIT License`）。

负例矩阵（临时副本上每次只改一个维度，其余部件与 ZIP 元数据原样保留，8 例全被抓）：

| 负例 | 期望命中 | 结果 |
| --- | --- | --- |
| 改坏第 1 节右边距 | 边距 | 命中 |
| 改正文文字 | 正文文字 | 命中 |
| 改标题行距 | 标题行距 | 命中（**本轮补齐后**才成立） |
| 改正文行距 | 正文第 1 段行距 | 命中 |
| 改 Normal 样式字号 | Normal 样式字号 | 命中 |
| 删一个 ZIP 部件 | 部件集合 | 命中 |
| 改固定元数据 | 作者 | 命中 |
| 破坏继承对照 | 继承对照 | 命中 |

统一校验 `tests/fixtures/synthetic/verify_fixtures.py` 已把本目录纳入：结构断言 1 组 / 10 份样本，清单一致性覆盖 5 个目录 / 51 份 docx。

Word 16 实测（build `16.0.20527`，`Office16`，只读打开、不保存）：

| 文件 | 段落 | 节 | 正文读数 | 边距 |
| --- | --- | --- | --- | --- |
| baseline_correct | 4 | 1 | 宋体 / Times New Roman / 12pt，行距规则 1（1.5 倍 18pt） | 左 3cm、其余 2.499cm |
| font_size_wrong | 4 | 1 | 宋体 / Times New Roman / **10.5pt** | 同上 |
| multisection_control | 8 | 2 | 同基线 | 同上 |
| style_override_control | 4 | 1 | P2/P3/P4 均 宋体 / TNR / 12pt（全部继承） | 同上 |
| style_override_wrong | 4 | 1 | **P3=楷体**，P2/P4 仍 宋体 / TNR / 12pt | 同上 |

即继承对照与唯一覆盖在 Word 对象模型里读出的差异，与 README 表格「document.xml 的 p[3]/r[1] 中文字体不匹配」一致。

CLI：10/10 样本只读加载成功，退出码 0，stderr 为空，输入 SHA-256 前后一致。加载器不检测格式，不把该结果当格式检测通过。

## 2026-11-10 重建路径实跑与 `--with-rebuild`

上一节写的「重建路径本轮未实跑」这一条本轮补掉了：装上 `python-docx==1.2.0`（Python 3.13.12）后实跑。

- **逐字节复现**：在临时目录里重建 10 份样本，与仓库版本的 SHA-256 **10/10 全部一致**。README 里「相同 Python/python-docx 环境下重复生成可逐字节核对」是可兑现的。
- **重建幂等**：原地跑 `generate.py --check`（重建 + 校验）后，10 份文件 SHA-256 与重建前完全相同，`git status` 无 `.docx` 变更。
- **`--with-rebuild`（新增，第四类判据）**：`verify_fixtures.py --with-rebuild` 在结构断言之外再对 10 份样本做逐部件重建校验。默认不开，保持零依赖；缺依赖时打印「重建校验已跳过：format-samples 需要 docx（pip install python-docx）」并以退出码 0 结束，**跳过不等于通过**。
- 判据验证：干净副本 0 问题 / 校验 10 份；负例 3/3 全抓（改正文文字、改字号、改边距）；**只改 ZIP 时间戳不报**（部件内容一致即通过），与「忽略 deflate 流与时间戳」的跨版本稳定口径一致。
- `--self-test` 从 8 条扩到 9 条，新增「缺依赖时重建校验跳过而非失败」。
- 为此把 `generate.py` 的重建逻辑抽成 `build_bytes(config)`，与其它生成器同名同义；`generate()` 改为调用它。

仍未覆盖：其他 Windows/Mac Word 版本、WPS、字体未安装或被替代、逐页视觉布局；以及**跨 Python / python-docx 版本的逐字节一致性**——上面 10/10 是在 Python 3.13.12 + python-docx 1.2.0 上得到的，换版本是否仍逐字节相同未验证（逐部件判据已为此做了时间戳与 deflate 解耦，但容器字节层面的结论只覆盖本机这一组版本）。
