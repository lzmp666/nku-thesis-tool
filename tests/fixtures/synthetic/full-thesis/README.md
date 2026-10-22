# 完整合成论文夹具

`synthetic_complete_thesis.docx` 是用于真实 Word 和引擎回归的人工构造样稿，不含学生、导师或真实研究数据。当前 Word 桌面版显示为 14 页、约 6,154 个字符；字体、分页和 Word 版本会影响实际页数。

论文内容覆盖中文与英文摘要、目录、六章正文、图、表、公式、参考文献、附录和致谢。正文预设三处错误，并用唯一文本标识：

- `【中文格式测试点】`：东亚字体设为黑体，预期宋体。
- `【字号格式测试点】`：字号设为 10.5 pt，预期 12 pt。
- `SYNTHETIC ENGLISH FORMAT CHECK`：西文字体设为 Arial，预期 Times New Roman。

分析器应只报这三处正文问题。表格单元格格式、图表题注、图片清晰度/布局、公式、页眉页脚均是当前未覆盖范围，报告必须明确列出，不得当作已通过。

可用 `python generate_full_thesis.py` 重建 DOCX。运行 `scripts/Test-WordInPlaceRepair.ps1` 可在 Microsoft Word 中验证全文段落映射、标记、批量修复和保存边界。重建后的哈希可能随文档属性或依赖版本变化。
