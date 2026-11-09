# 文档级版式合成样本

Issue #112 的样本集：纸张、页边距、页眉页脚距离、装订线、页眉页脚、页码域、页码格式、目录、疑似空白页。全部人工构造，不含真实论文内容。

## 生成

```sh
python tests/fixtures/synthetic/document/generate.py          # 重新生成
python tests/fixtures/synthetic/document/generate.py --check  # 只比对，不写盘
```

只用 Python 标准库。ZIP 条目顺序与时间戳固定，同一 Python 版本下重复生成可逐字节核对。

## 场景

每个文件只偏离一个维度，预期见 `manifest.json`。「规范」指 `rules/nku-graduate/general.yaml` 有明文条款的项，「惯例」指规范没写、按排版习惯报的项。

| 文件名 | 偏差 | 预期 |
| --- | --- | --- |
| document_baseline_correct | 无 | 0 条 |
| document_paper_wrong | Letter 纸张 | D1 一条 |
| document_margin_wrong | 上下 2.5cm | D2 两条（左右不该跟着报） |
| document_hf_distance_wrong | 页眉页脚 1.27cm | D3 两条 |
| document_gutter_wrong | 装订线 1cm | D4 一条 |
| document_no_header | 没有任何页眉引用 | D6 一条 |
| document_empty_header | 页眉部件为空 | D6 一条 |
| document_footer_text | 页脚手打数字 | D7 一条 |
| document_even_odd | 开启奇偶页不同 | D5 一条 |
| document_toc_manual | 目录是手打的 | D9 一条 |
| document_toc_auto | 目录是 TOC 域 | 0 条 |
| document_blank_page | 空段落带分页符 | D10 一条（惯例） |
| document_sections_same_numbering | 两节都是 decimal | D8 一条 |
| document_sections_mixed_numbering | 罗马 + 阿拉伯 | 0 条 |

## 两条按全文判断的规则，以及为什么

`D6`（页眉）和 `D7`（页码）**不逐节报**，只在全文缺失时报一条。这不是偷懒，是实测逼出来的：

用 Word 16.0 生成一份合规文档后解包，`sectPr` 里是这样的：

```xml
<w:sectPr>
  <w:headerReference w:type="even" r:id="rId6"/>
  <w:headerReference w:type="default" r:id="rId7"/>
  <w:footerReference w:type="even" r:id="rId8"/>
  <w:footerReference w:type="default" r:id="rId9"/>
  <w:headerReference w:type="first" r:id="rId10"/>
  <w:footerReference w:type="first" r:id="rId11"/>
  …
</w:sectPr>
```

Word 会给每一节都建好 default / even / first 三组页眉页脚引用，内容是空的。而规范说的是「篇眉从中文摘要开始」——封面本来就没有页眉。按节判「页眉为空」，每篇合格论文都会被报一遍，用户第一次用就会觉得这工具在乱叫。

「哪一节是封面」在 XML 层面判不出来（标题文字千变万化），所以只报全文缺失。某一节为空的情况放在报告开头的读数里，用户自己看得到。

## Word 存出来的数字不是教科书上的数字

| 项 | 教科书值 | Word 16 实际存的 |
| --- | --- | --- |
| A4 纸张 | 11906 × 16838 twips | **11907 × 16839** |
| 3.8cm 页边距 | 2155 twips | **2154** |
| 页眉页脚 3.0cm | 1701 twips | 1701 |

检查用 0.1cm 容差吸收这种取整差，不做精确相等比较。写死 `== 11906` 的检查会把自己的合规样本判成不合规。

## 未验证的部分

**样本没有在 Word 里打开过**，脚本只能保证 ZIP 与 XML 结构合法。

但**真实 Word 文档实测过**：2026-11-09 用 Word 16.0（COM 自动化）生成了 A4、3.8/3.2cm、页眉页脚 3.0cm、页眉「第一章 引言」、页脚 PAGE 域的文档，`analyze-document` 输出 0 条问题，读数与设置一致。上面两条「按全文判断」和「容差」的口径都是在那次实测后改的。

自动修复（只动 `sectPr`）不在本档范围内，属 11/29 之后的部分。
