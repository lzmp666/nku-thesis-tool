# 公式排版合成样本

Issue #114 第一档（只查不改）的样本集。全部人工构造，公式都是简单占位表达式，不含真实论文内容。

## 生成

```sh
python tests/fixtures/synthetic/formula/generate.py          # 重新生成
python tests/fixtures/synthetic/formula/generate.py --check  # 只比对，不写盘
```

只用 Python 标准库，不需要 python-docx 或 lxml。ZIP 条目顺序与时间戳固定，同一 Python 版本下重复生成可逐字节核对。

包骨架（`[Content_Types].xml`、`_rels/.rels`、`word/_rels/document.xml.rels`、`styles.xml`、`settings.xml`）照 `demo/clean_sample.docx` 的写法内联在脚本里，避免手写包结构出错。`sectPr` 用规范值（A4、上下 3.8cm、左右 3.2cm），所以版心宽固定为 8278 twips，右制表位以此为准。

## 场景

每个文件只偏离一个维度，预期见 `manifest.json`。表里的「规范」指 `rules/nku-graduate/general.yaml` 有明文条款的项，「惯例」指规范没写、按排版习惯报的项。

| 文件名（均为 .docx） | 偏差 | 预期 |
| --- | --- | --- |
| formula_baseline_correct | 无 | 2 个独立公式、2 个行内公式，0 条问题 |
| formula_not_centered | 公式段左对齐 | 惯例 1 条（F1 未居中） |
| formula_number_not_right_aligned | 无右制表位，段落居中而非右对齐 | 规范 1 条（F2 式号未右对齐） |
| formula_number_format_wrong | 式号写裸数字 3、宋体、小四 | 规范 3 条（F3 括号 / 字体 / 字号） |
| formula_spacing_wrong | 公式段段前段后各 6 磅 | 规范 1 条（F5 间距） |
| formula_table_row_correct | 无，两列表格排版 | 0 条，且识别为表格排版 |
| formula_word_native_correct | 无，Word 16 原生写法（式号在公式里） | 0 条，式号要能认出来 |
| formula_word_native_number_not_aligned | Word 原生写法但左对齐、无右制表位 | 规范 1 条（F2）+ 惯例 1 条（F1） |
| formula_inline_only | 无，只有行内公式 | 独立公式 0 个、行内 2 个、0 条问题 |

## Word 原生写法与合成样本的差别

`formula_word_native_*` 两个样本是照 2026-11-09 在 Word 16.0 上实测的输出构造的，结构与其余样本不同，差别值得记住：

```xml
<w:pPr><w:tabs><w:tab w:val="right" w:pos="8278"/></w:tabs><w:jc w:val="center"/></w:pPr>
<m:oMathPara><m:oMath>
  <m:r><w:rPr><w:rFonts w:ascii="Cambria Math" …/><w:sz w:val="21"/></w:rPr><m:t>E=mc^2</m:t></m:r>
  <m:r><w:rPr><w:rFonts w:ascii="Cambria Math" …/><w:sz w:val="21"/></w:rPr><w:tab/><m:t>(1)</m:t></m:r>
</m:oMath></m:oMathPara>
```

- **式号在公式对象内部，是 `m:t` 不是 `w:t`**。只读 `w:t` 的实现会认不出式号，于是「式号未右对齐」「式号写法不对」两项永远不检查，报告却显示 0 条。
- **没有 `m:oMathParaPr`**，居中靠 `w:pPr/jc`。所以两种居中写法都得认。
- **式号的字体是 `Cambria Math`**，因为它在公式里由公式引擎决定，不是用户选的。所以式号在公式内时只查字号不查字体——报字体等于让所有用 Word 公式编辑器的人都背着一条改不掉的错。

## 两个必须保留的反例

这两个不在样本文件里，写在 `tests/ThesisEngine.Tests/FormulaAnalyzerTests.cs` 里现场造：

1. **纯 OMML 段没有 `w:r`、没有 `w:t`**。公式必须从 `m:oMathPara` / `m:oMath` 反向枚举父段落，不能挂在段落文本或 run 遍历下面，否则这类段落整段跳过，报告 0 条而实际全不合规。
2. **句子中间的公式**。段落里既有正文文字又有公式时，它是行内公式，不能被套上「未居中 / 未右对齐」。

## 未验证的部分

**样本文件本身没有在 Word 里打开过。** 脚本只能保证 ZIP 与 XML 结构合法。第一档只查不改，产物就是原始文档，不涉及写入。

**但 Word 原生结构实测过**：2026-11-09 用 Word 16.0（COM 自动化）生成了两份真实公式文档，`analyze-formula` 的结论与人工预期一致——居中且带右制表位的那份 0 条，左对齐无制表位的那份报出 F2 与 F1。`formula_word_native_*` 两个样本就是按那次实测的 XML 结构固化的，因为 Word 生成的文件含 rsid 和时间戳、无法用脚本确定性重现，所以不直接提交原件。

公式渲染未验证：`m:t` 里的文本只用于报告预览，不参与判断，符号在 Word 里是否显示正常不在本样本的验证范围。
