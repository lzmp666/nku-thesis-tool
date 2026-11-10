# 文档级版式合成样本

Issue #112 的样本集：纸张、页边距、页眉页脚距离、装订线、页眉页脚、页码域、页码格式、目录、
空白页，以及第二档的 R2 篇眉 / R3 页码字体字号与居中。全部人工构造，不含真实论文内容。

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
| document_blank_page | 空段落带分页符 | D10 一条（惯例；Word 实测只多 1 页，第 2 页仍有正文） |
| document_sections_same_numbering | 两节都是 decimal | D8 一条 |
| document_sections_mixed_numbering | 罗马 + 阿拉伯 | 0 条 |
| document_header_font_wrong | 篇眉黑体 | D11 一条 |
| document_header_size_wrong | 篇眉 12 磅 | D11 一条 |
| document_header_left_aligned | 篇眉左对齐 | D11 一条 |
| document_footer_font_wrong | 页码黑体 | D12 一条 |
| document_footer_size_wrong | 页码 12 磅 | D12 一条 |
| document_footer_left_aligned | 页码左对齐 | D12 一条 |
| document_abstract_header_correct | 全西文篇眉用 Times New Roman | 0 条 |
| document_abstract_header_wrong | 全西文篇眉用 Arial | D11 一条 |
| document_hf_style_inherited | 字体字号居中写在 styles.xml 的样式里 | 0 条 |
| document_no_styles_part | 无 styles.xml 且无显式 rPr / jc | 0 条，但完整判定 0/2，不能当合规 |
| document_no_styles_explicit | 无 styles.xml 但显式写了格式 | 0 条 |
| document_font_aliases | SimSun / TimesNewRoman 别名写法 | 0 条 |
| document_true_blank_page | 连续两个空段落各带分页符 | D10 一条（Word 实测 3 页、第 2 页整页无正文） |

第二档那批要注意 `document_no_styles_part`：它跑出 0 条，是因为三个维度都**判不出来**，
不是因为查过没问题。报告头部会显示「完整判定 0 / 2」，末尾「本次未检查」逐条列出原因。
这条是专门用来守「0 条要能分清查了没问题和压根没查」的规矩。

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

## Word 实机核验

2026-11-09 用 Word 16.0（COM 自动化）生成了 A4、3.8/3.2cm、页眉页脚 3.0cm、
页眉「第一章 引言」、页脚 PAGE 域的文档，`analyze-document` 输出 0 条问题，读数与设置一致。
上面「按全文判断」和「容差」两条口径都是在那次实测后改的。

2026-11-10 又把这批合成样本本身放进了 Word 16.0：修复件页边距读回 3.799/3.2/3.0cm、
纸张 21.001×29.7cm，修复前后 `Content.Text` 逐字相同；页脚是 `Field.Type=33` 的 PAGE 域，
域结果宋体 10.5 磅、居中；空白页两个样本分别实测 2 页与 3 页，后者第 2 页无正文。
原始读数、逐页文本与四张截图见 `docs/research/word-112-acceptance.md`。

正因为这次实测发现「单个空段落带分页符」并不会产生整页空白，才补了
`document_true_blank_page.docx`，并把 D10 拆成两种措辞。
