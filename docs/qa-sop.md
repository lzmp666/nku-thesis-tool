# QA SOP
synthetic/：自行构造最小 Word 文档，单一错误到组合错误，包含样式/直接格式、继承、东亚字体、多节边距等；样本生成方法和 Word 版本一并记录。
real-anonymized/：先获得使用许可，再脱敏，再复核；原始文件留在受控存储，不进入 Git。见该目录 README。
goldens/：人工核对的标准记录，不是简单保存引擎输出；采用 template.yaml，稳定记录规则、节点、expected/actual、修复等级与未覆盖项。
回归：restore/build/test；检测正确样本无误报，错误样本有定位；修复后检查内容/关系/样式/节未受损，二次 diagnostic=0、patch=0。suggest/manual 未处理项必须保留，不得假报合规。
需要时比较规范化 XML（命名空间、属性顺序等）并按部件说明合理差异，不用整个 docx 哈希。
Bug：用 Bug 模板，写最小复现、命令、版本、实际/期望、脱敏样本与 golden；截图必须脱敏，不能仅写“不能用”。无法分享真实文档时提供 synthetic 重现。
合并前至少一名非作者 Review；正式规则由负责人审核，未支持项明确记录。
