# R7–R10 标题层级格式检查样本

依据 `rules/nku-graduate/general.yaml` 4.7 表4.6 的标题规则（来源 S4 2025 版 / S1 2026 版，
规则状态 draft），供 `analyze-headings` 命令与 `HeadingFormatAnalyzer` 单元测试使用。

## 标题识别口径

标题级别只按 `w:outlineLvl` 或 `HeadingN` /「标题 N」样式（沿 `basedOn` 链）识别，
不靠字体特征猜测。本批样本的 `styles.xml` 里 Heading1–4 分别带 outlineLvl 0/1/2/3，
段落用 `w:pStyle` 指向它们。章起点还需文本像章标题；检查到参考文献/附录/致谢标题前止。

## 规范值与问题码

| 级别 | Rule | 字体 | 字号 | 加粗 | 行距 | 段前 | 段后 | 码段 |
|---|---|---|---|---|---|---|---|---|
| 章 | R7 | 黑体 | 16 磅 | 是 | 单倍 | 24 磅 | 18 磅 | H1–H6 |
| 一级 | R8 | 黑体 | 14 磅 | 是 | 单倍 | 24 磅 | 6 磅 | H7–H12 |
| 二级 | R9 | 黑体 | 13 磅 | 不判 | 单倍 | 12 磅 | 6 磅 | H13–H17 |
| 三级 | R10 | 黑体 | 12 磅 | 不判 | 单倍 | 12 磅 | 6 磅 | H18–H22 |

码段用每级 CodeStart + 维度偏移映射，H1–H22 连续，避免字符串拼接产生 H111。
二级/三级原文未写加粗（general.yaml 明确「不得默认为加粗」），不判该维度。

## 未判定维度

- 对齐：模式一居左、模式二居中，但 S4 与 S5/S7 表述冲突（`docs/rule-notes/headings.md` T2/C3）。
- 序号与题名间空一个汉字：宽度语义未定义（headings.md RFC-3）。
- 以上统一列「未判定」，不伪造硬规则；这些维度样本也不制造差异。

## 样本

见 `manifest.json`。只用 Python 标准库生成，ZIP 条目顺序与时间戳固定，可逐字节重建核对。

```bash
python generate.py           # 重新生成
python generate.py --check   # 比对，不写盘
```
