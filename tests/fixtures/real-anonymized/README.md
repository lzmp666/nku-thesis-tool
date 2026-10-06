# 真实样本授权 脱敏与复核

当前无真实样本。本文件是检查表及空记录模板，不表示任何真实文档已获许可或通过复核。原件与身份映射仅存受控本地或批准的安全存储，不进入 Git；Private 仓库也不例外。

## 授权与另存

先取得所有者/作者的明确授权，记录允许用途、接收方、许可期限和是否允许进入本仓库。书面授权原件和真实身份映射保留在受控存储；仓库只记录匿名 sampleId、许可范围摘要及受控记录编号，不上传含身份信息的授权书。没有仓库提交许可时，即使完成脱敏也不得提交。

在副本上另存脱敏文件，保护原件。不以“隐藏修订”“裁剪显示区域”或隐藏文本代替清除敏感内容。

## 脱敏检查表

每项填写通过/不适用/未完成及理由；未完成项阻止提交。

- [ ] 正文、表格、目录、文本框、字段结果、隐藏文字、超链接显示文本：去除姓名、学号、联系方式、导师信息、致谢、签名及能重新识别作者的内容。
- [ ] 元数据：检查 docProps/core.xml、app.xml、custom.xml（若存在），清除作者、公司、最后修改人、模板路径和自定义属性；不只改文件名。
- [ ] 批注：检查 word/comments*.xml、people 相关部件（若存在）及 document/header/footer/footnotes/endnotes 中的批注标记、作者和内容。删除相关部件时同步清理关系、内容类型与锚点。
- [ ] 修订：在 Word 显示全部修订，按授权范围接受/拒绝并复核最终文本。修订位于 document.xml、header/footer、footnotes/endnotes 等内容部件中的 w:ins、w:del、w:delText、w:moveFrom/w:moveTo 以及 *PrChange 等节点；检查 w:author、w:initials、日期和被删除内容。不存在通用的 word/trackChanges 文件夹；关闭 settings.xml 中的 w:trackRevisions 只停止新修订，不清除已有记录。
- [ ] 页眉页脚、脚注尾注：逐个检查 word/header*.xml、footer*.xml、footnotes.xml、endnotes.xml 的文字、图片、域和关系。
- [ ] 图片：检查 word/media/ 的实际二进制内容及替代文字/标题。删除隐藏原图和裁剪区域，不只覆盖形状或打上遮挡层；重新导出真正脱敏图片再替换，检查人脸、签名、证件与其他识别信息。
- [ ] 嵌入对象：检查 word/embeddings/ 和 OLE/附件关系，移除含原始数据的表格、源文件及缓存预览；不能仅改可见图标。
- [ ] 外链与关系：遍历包内所有 .rels 的 TargetMode=External，检查本地路径、服务器地址、超链接、链接图片和附加模板。按许可范围移除/替换敏感目标并确认文档结构仍有效。
- [ ] 自定义 XML 与数据绑定：检查 customXml/、对应关系和 w:dataBinding；清除姓名/学号等隐藏数据及失效绑定，不让它们恢复到可见内容中。
- [ ] Word Document Inspector：在指定 Word 版本执行“文件 → 信息 → 检查问题 → 检查文档”，检查全部适用类别；处理后再次检查。Inspector 与逐部件检查共同使用，不以一次扫描代替上述检查。
- [ ] 第二人复核：由脱敏执行者以外的人检查可见内容、解包部件、元数据和关系，确认匿名标识、授权范围与提交内容一致；记录版本、日期及未检查范围。
- [ ] 提交前复核：只提交批准的匿名文件和说明；不提交原件、真实身份映射、敏感授权书、临时解包目录或带本机敏感路径的日志。

## 仓库记录模板

复制下表填写；示例占位符不是已执行记录。采用团队批准的非敏感成员标识，授权详情和身份映射不入仓库。

| sampleId | 匿名文件名 | 许可范围摘要 | 受控授权记录编号 | 脱敏执行者标识 | 第二复核者标识 | Word 版本/系统 | 日期 | 检查表与未检查范围 | 结果 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 待填写 | 待填写 | 待确认 | 待确认 | 待填写 | 待填写 | 待实测 | 待执行 | 待执行 | pending |

结果只能在授权、逐项检查和第二人复核均完成后填通过；不适用项写明理由。无真实样本时保留 pending，不伪造签字或日期。

## 官方操作依据

查阅日期 2026-10-06：
- Microsoft Open XML 修订处理例子（ins/del/pPrChange 等，不是固定 trackChanges 部件）：https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document
- Microsoft Document Inspector：https://support.microsoft.com/en-us/office/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks-356b7b5d-77af-44fe-a07f-9aa4d085966f
