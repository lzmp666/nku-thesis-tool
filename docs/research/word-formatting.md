# Word 有效格式继承与覆盖调研

本记录交付 Issue #8 的解析依据和最小例子，不实现引擎，不把合成字体和字号当作学校规范。版本 0.2，2026-10-06 自审；负责人及非作者审核待完成。

## 属性来源与主题解析

对普通正文的字体、字号等非 toggle 属性，先取得 docDefaults，再解析段落样式及其 basedOn 链、字符样式及其 basedOn 链，最后应用 run 直接属性。每个属性独立合并：run 只写字号时仍需继承字体。完整层级还包含表格和编号样式，不能把正文的简化路径用于所有节点。[R1][R2]

主题不是位于默认值和样式之间的统一覆盖层。asciiTheme/hAnsiTheme/eastAsiaTheme/cstheme 是各格式来源中的字体引用，需结合主题部件、脚本和语言解析；跨层级覆盖与同一 rFonts 内的属性冲突须分开处理。Microsoft 的实现说明明确区分这两种情况；同层同时出现字体名和主题引用的冲突不在当前例子的支持范围内，不推断统一的名称优先规则。[R3]

basedOn 的 val 引用 styleId，不是字体名称。按父到子解析同类型样式；父样式不存在或类型不匹配时不能凭空补父属性。实现应检测循环和重复 ID，报告无法判断，不让递归无限运行。字符样式不用于覆盖段落行距或节边距。[R2]

字体选择需区分 ASCII、hAnsi、eastAsia、cs 及语言/脚本信息。w:ascii 改变英文字符的字体，不等于中文字体改变；字体名称已知也不证明本机安装或实际渲染字体已知。[R3][R4]

## 最小例子与对应节点

以下片段共用 w 命名空间 http://schemas.openxmlformats.org/wordprocessingml/2006/main；都是隔离的合成输入。配套 document 与 styles XML 在 tests/fixtures/synthetic/research/，只是部件示例，不是独立可打开的 docx。

### 1 默认值被英文直接字体覆盖

styles.xml 的 docDefaults/rPrDefault/rPr/rFonts 为 ascii=TestFontA；document.xml 的 p[1]/r[1]/rPr/rFonts 为 ascii=TestFontB，文字为 ABC。预期 ASCII 字体 TestFontB，不能把该断言推广到中文。[R4]

### 2 段落样式父链继承

styles.xml：Parent 为 paragraph，rPr/sz=24；Child 为 paragraph，basedOn=Parent，自身不写 sz。document.xml 的 p[2]/pPr/pStyle=Child，run 不写 sz。预期字号 12pt。查找位置是 run，来源是 Parent；不能因直接 sz 缺失报未知。[R2][R5]

### 3 字符样式覆盖段落样式

Para 样式 rPr/sz=24；Char 样式 type=character，rPr/sz=21。p[3] 引用 Para，r[1]/rPr/rStyle 引用 Char，run 无直接 sz。预期 10.5pt。[R1][R5]

### 4 直接字号覆盖字符样式

p[4] 与例子 3 一致，但 r[1]/rPr/sz=22。预期 11pt。24/21/22 都是半点整数，不能四舍五入为整点，也不能把 sz=0 当有效字号。[R5]

### 5 跨层字体名覆盖主题引用

ParaTheme 样式 rPr/rFonts 只写 asciiTheme=minorHAnsi；p[5]/r[1] 直接 rFonts 只写 ascii=TestFontB，文字 ABC。预期 ASCII 字体 TestFontB。这是两个层级，不是同一个 rFonts 中同时写两个冲突属性。[R3]

### 6 只有主题引用时先解析主题

run 只有 asciiTheme=minorHAnsi，主题 fontScheme/minorFont/latin/@typeface=TestFontA：预期 ASCII 字体 TestFontA。若主题部件不存在或映射不可解析，结果为未覆盖/无法判断，不把 minorHAnsi 字符串当作字体名；东亚主题还需语言和脚本映射。[R3][R4]

## 反例与实现边界

- run 只写 sz=22 时，不能丢弃来自父样式的字体；格式合并按属性进行。
- 一个 run 含 ABC中文 时，不能用 ascii 字体代表全部字符；混合多个 run 也需逐节点定位。
- 在 rPr 外写 rFonts 是结构无效输入，不能承诺 Word 必定静默忽略或回退默认值。
- 缺失父样式、样式循环、主题冲突、复杂脚本字体、表格条件样式、编号、toggle 属性（如样式链中的 b/i）、兼容模式的实际布局均需单独验证。本文不证明这些范围已实现，不把未知项当匹配。
- 规范合并依据不是 Word 实测记录；当前 Core 仅支持加载。后续有效格式解析须按例子补执行测试，再核对 Word 版本、字体安装和 UI 结果。

## 官方依据

以下按概念转述，查阅日 2026-10-06。不存在学校正式规则声明。

- R1 Microsoft 维护的 Simplify-Docx 对 pStyle 层级的说明（document defaults、table、numbering、paragraph、character、direct）：https://github.com/microsoft/Simplify-Docx/blob/master/src/simplify_docx/utils/paragrapy_style.py 。该项目实现不是本工具的完整解析器。
- R2 Microsoft Learn BasedOn，Open XML SDK 3.0.1，Remarks 中父 styleId、同类型限制和父链例子：https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.basedon?view=openxml-3.0.1
- R3 Microsoft MS-OI29500 §2.1.88 / ISO Part 1 §17.3.2.26 rFonts，脚本选择及不同层级的字体名/主题覆盖：https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/aef3c9a6-5d6c-434b-90b7-85e761fd8e62
- R4 Microsoft Learn RunFonts，Open XML SDK 3.0.1，Remarks 与 Parent Elements：https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts?view=openxml-3.0.1
- R5 Microsoft Learn FontSize，Open XML SDK 3.0.1，半点数和未设置时的继承：https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fontsize?view=openxml-3.0.1
