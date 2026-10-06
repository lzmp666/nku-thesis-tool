# Word 格式继承与覆盖调研（OOXML）

## 背景与目标
研究 Word 在排版时的有效格式优先级，为后续解析引擎提供理论基础。

## OOXML 优先关系
在 OOXML (WordprocessingML) 中，样式与格式的应用优先级从低到高依次为：
1. 文档默认值（docDefaults）
2. 主题（Theme）样式
3. 段落样式/字符样式
4. 直接格式（Direct Formatting）

**结论**：直接格式 > 字符/段落样式 > 主题 > 默认值。

## 最小例子

### 例子 1：默认字体 vs 直接格式
- 输入：文档默认字体为“宋体”，某段落直接格式设为“Arial”。
- 节点：`<w:rPr><w:rFonts w:ascii="Arial" .../>`
- 预期结果：Arial 生效。

### 例子 2：段落样式 vs 字符样式
- 输入：段落样式设为“宋体”，该段落内的某几个字使用字符样式“黑体”。
- 节点：`<w:pStyle w:val="宋体"/>` + `<w:rStyle w:val="黑体"/>`
- 预期结果：黑体生效。

### 例子 3：主题字体 vs 直接格式
- 输入：段落使用了“主题字体（等线）”，但该段落直接指定了“Times New Roman”。
- 节点：`<w:rFonts w:asciiTheme="minorHAnsi"/>` 加上 `<w:rFonts w:ascii="Times New Roman"/>`
- 预期结果：Times New Roman 生效。

### 例子 4：半点字号（10.5pt）的覆盖
- 输入：样式规定字号为 12pt，直接格式规定字号为 10.5pt。
- 节点：`<w:sz w:val="21"/>`（OOXML 中半点乘以2，21 代表 10.5pt）
- 预期结果：10.5pt 生效，引擎需正确识别半点。

## 反例
- **反例 1**：如果将 `<w:rFonts>` 写在 `<w:rPr>` 之外，该格式将被忽略，导致直接格式失效，继承默认值（错误）。
- **反例 2**：混合 Run 中，中文和英文字体设置不一致，如果引擎只读取第一个 Run 的字体，会导致错误诊断。

## Microsoft 官方资料与对应节点
- [ECMA-376 Part 1 - 17.7.2 样式继承](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- 节点说明：`w:docDefaults`, `w:style`, `w:rPr`, `w:pPr`

## 未知/不支持项
- 部分复杂的主题字体（如东亚主题字体与西文主题字体的级联）尚未完全覆盖，需进一步调研。
- 部分旧版 Word 的兼容模式（Compatibility Mode）可能存在优先级差异，暂不支持。