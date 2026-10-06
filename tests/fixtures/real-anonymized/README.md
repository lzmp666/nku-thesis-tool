# 真实样本脱敏与复核说明

## ⚠️ 核心原则
**严禁将包含原始敏感信息的论文直接提交到本仓库。** 即使是 Private 仓库，也必须进行严格的脱敏处理。

## 脱敏检查表 (Document Inspector & OOXML 部件检查)
在处理真实文件时，必须经过以下检查步骤：
- [ ] **元数据 (Metadata)**：检查并清除 `docProps/core.xml` 和 `docProps/app.xml` 中的作者、公司、最后修改人等信息。
- [ ] **批注/修订 (Comments & Revisions)**：检查 `word/comments.xml` 和 `word/trackChanges`，确保所有批注和修订均已清除或匿名化。
- [ ] **页眉页脚 (Headers & Footers)**：检查 `word/header*.xml` 和 `word/footer*.xml`，确保没有包含敏感信息的文本或图片。
- [ ] **图片 (Images)**：检查 `word/media/` 中的图片，确保已裁剪或打码任何敏感内容（如人脸、身份证号）。
- [ ] **嵌入对象 (Embedded Objects)**：检查 `word/embeddings/`，确保没有嵌入包含敏感数据的 Excel 或源文件。
- [ ] **外链 (External Links)**：检查文档中的超链接，确保没有指向本地文件或私有服务器的链接。
- [ ] **第二人复核 (Second-person Review)**：脱敏完成后，必须由另一位成员进行检查并签字确认。

## 授权与复核记录表
| 文件名 | 原始来源 | 授权状态 | 脱敏执行人 | 复核人 | 复核日期 | 结果 |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| (示例) sample1.docx | 某某大学 | 已获得书面授权 | 张三 | 李四 | 2026-10-06 | 通过 |
