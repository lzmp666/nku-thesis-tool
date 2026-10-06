# 合成格式样本（tests/fixtures/synthetic/）

为 issue #5 建立的可分享最小格式样本集。用于格式检查器的正确 / 错误用例回归测试。

**重要声明：**

- 所有样本均为**自建合成文档**，文本为占位性质的合成文字，**不复制任何真实论文内容**；
- 每个错误样本**只引入一个可说明的格式错误**，其余属性保持基线正确，便于检查器定位与断言；
- 本目录下的样本可自由分享，无版权与隐私风险。

## 基线格式约定（结构样本阶段假设值）

> **注意：** 以下数值为结构样本阶段的**假设基线**。格式期望依赖规范审核（issue #5 前置依赖），
> 规范确认后只需修改 `generate.py` 顶部的 `BASE` 字典并重新生成，样本与本文档数值同步更新。

| 属性 | 约定值 |
|---|---|
| 中文正文字体 | 宋体 |
| 西文字体 | Times New Roman |
| 正文字号 | 小四（12pt） |
| 一级标题 | 黑体，三号（16pt），居中，不加粗 |
| 正文行距 | 1.5 倍 |
| 页边距 | 上 2.5cm / 下 2.5cm / 左 3.0cm / 右 2.5cm |

## 样本清单

| # | 文件 | 场景 | 注入的错误（唯一） | 人工预期 |
|---|---|---|---|---|
| 1 | `baseline_correct.docx` | 全部属性符合基线的对照组 | 无 | 检查器不报告任何格式问题 |
| 2 | `font_family_cn_wrong.docx` | 中文正文字体错误 | 宋体 → 楷体 | 检查器报告正文中文字体为楷体；Word 中选中正文，字体框显示“楷体” |
| 3 | `font_ascii_wrong.docx` | 西文字体错误 | Times New Roman → Arial | 检查器报告西文字体为 Arial；Word 中西文与数字（"The quick brown fox…"、数字）显示为 Arial |
| 4 | `font_size_wrong.docx` | 正文字号错误 | 小四 12pt → 五号 10.5pt | 检查器报告正文字号为 10.5pt；Word 字号框显示 10.5 |
| 5 | `line_spacing_wrong.docx` | 正文行距错误 | 1.5 倍 → 单倍 | 检查器报告行距为单倍；Word“段落”对话框显示单倍行距 |
| 6 | `margin_wrong.docx` | 页边距错误 | 左边距 3.0cm → 2.0cm | 检查器报告第 1 节左边距 2.0cm；Word“页面设置”可见左 2.0cm |
| 7 | `style_override_control.docx` | 样式统一、无直接覆盖（对照组） | 无 | 检查器不报告错误；验证检测逻辑基于有效字体（含样式继承），不误报 |
| 8 | `style_override_wrong.docx` | 样式正确但被直接格式覆盖 | 第 2 段 run 直接格式把中文字体改为楷体（Normal 样式本身仍为宋体） | 检查器报告第 2 段中文字体为楷体；在 Word 中修改 Normal 样式字体，该段外观不变（证明是 run 级覆盖） |
| 9 | `multisection_control.docx` | 多节文档（2 节），两节均正确 | 无 | 检查器不报告错误；验证检测器遍历所有节而非只看第 1 节 |
| 10 | `multisection_margin_wrong.docx` | 多节文档，仅第 2 节错误 | 第 2 节左边距 3.0cm → 2.5cm（第 1 节正确） | 检查器仅报告第 2 节左边距 2.5cm，第 1 节不报错 |

共 10 个样本（要求 ≥ 8），其中 4 个正确对照、6 个单错误样本，覆盖：字体（中文/西文）、字号、行距、页边距、样式覆盖、多节。

## 生成方式与步骤

### 方式 A：脚本生成（推荐，本仓库采用）

样本由 `generate.py` 生成，可重复再生成（内容确定性一致）：

```bash
# Python 3.9+，依赖：python-docx >= 1.0
pip install python-docx

cd tests/fixtures/synthetic/
python generate.py           # 生成全部 10 个样本
python generate.py --check   # 生成后自动重新读取并校验关键属性
```

- 生成工具：Python 3.13 + python-docx 1.2.0
- 产物为标准 OOXML（ECMA-376）`.docx`，未使用任何非标扩展

### 方式 B：在 Word 中手工重建（等价步骤）

若需用 Word 手工重建任一样本（以 `margin_wrong.docx` 为例）：

1. 新建空白文档；
2. “布局 → 页面设置 → 页边距”：上 2.5 / 下 2.5 / 左 **2.0**（错误值）/ 右 2.5；
3. 选中正文，“开始 → 字体”设为宋体 + Times New Roman，小四；段落行距 1.5 倍；
4. 输入合成占位文本（见 `generate.py` 中 `SYNTH_PARAS`），标题用黑体三号居中；
5. 另存为 `.docx`。

其余样本同法，仅把对应属性换成表中“注入的错误”值。

## Word 版本兼容性

- 文件为标准 OOXML 格式，目标兼容 **Microsoft Word 2010 / 2013 / 2016 / 2019 / Microsoft 365（Windows）**、**Word for Mac 2016+**、**WPS Office**；
- 生成时未使用主题字体（theme font）引用，run 级字体/字号均显式写入，避免不同版本主题差异导致的渲染漂移；
- 建议人工复核环境：Microsoft Word 2016 或更高版本（Windows）；
- 人工复核要点：打开文件无修复提示；按上表“人工预期”逐项在 Word UI（字体框 / 段落对话框 / 页面设置）中确认。

## 人工复核清单（建议执行后在此打勾）

- [ ] 10 个文件在 Word 2016+ 中均能正常打开，无“文件损坏/修复”提示
- [ ] `baseline_correct.docx` / `style_override_control.docx` / `multisection_control.docx` 目测与基线一致
- [ ] 各错误样本在 Word UI 中确认唯一错误项与表中一致，其余属性正常
- [ ] `style_override_wrong.docx`：修改 Normal 样式字体后第 2 段外观不变
- [ ] `multisection_margin_wrong.docx`：光标置于第 2 节，页面设置显示左 2.5cm；第 1 节显示 3.0cm

## 目录内文件

```
tests/fixtures/synthetic/
├── README.md                      # 本文档
├── generate.py                    # 生成 + 自校验脚本
├── baseline_correct.docx
├── font_family_cn_wrong.docx
├── font_ascii_wrong.docx
├── font_size_wrong.docx
├── line_spacing_wrong.docx
├── margin_wrong.docx
├── style_override_control.docx
├── style_override_wrong.docx
├── multisection_control.docx
└── multisection_margin_wrong.docx
```

## 维护说明

- 基线数值变更（规范审核后）：修改 `generate.py` 的 `BASE` 字典 → `python generate.py --check` → 同步更新本文档表格；
- 新增维度样本（如段前段后间距、页码、脚注）：在 `generate.py` 注册表中追加条目并更新本文档。
