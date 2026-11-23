# R12 图题/表题格式检查样本

依据 `rules/nku-graduate/general.yaml` 4.7 表4.5 的图题/表题规则（来源 S1 2025 版，规则状态 draft），
供 `analyze-captions` 命令与 `CaptionFormatAnalyzer` 单元测试使用。

## 检测对象与口径

- 题注用 `FigureTableParagraphs.Classify` 的文本锚点识别（「图 1-1 / 表 2-3 / 续表 / Figure / Table」
  带编号），不依赖图片/表格对象的视觉识别。general.yaml 把本组标为 `v1_check: false`
  （对象识别未决），本检查只判题注段落自身的纯文本格式，规避该未决点。
- 题注相对图表对象的位置（图题在图下、表题在表上）、序题间空一个汉字（RFC-3）、
  R13 表达式编号（Times New Roman 10.5 磅右对齐，依赖公式对象/渲染）均列未判定。

## 规范值与问题码

| 部件 | Rule | 字体 | 字号 | 对齐 | 行距 | 段前 | 段后 | 码段 |
|---|---|---|---|---|---|---|---|---|
| 图题（图下） | R12 | 宋体 | 10.5 磅 | 居中 | 单倍 | 6 磅 | 12 磅 | K1–K6 |
| 表题（表上） | R12 | 宋体 | 10.5 磅 | 居中 | 单倍 | 6 磅 | 6 磅 | K7–K12 |

码段用每类 CodeStart + 维度偏移映射，K1–K12 连续。字号 OOXML 用半点（val=21 即 10.5 磅）。

## 样本

见 `manifest.json`。只用 Python 标准库生成，ZIP 条目顺序与时间戳固定，可逐字节重建核对。

```bash
python generate.py           # 重新生成
python generate.py --check   # 比对，不写盘
```
