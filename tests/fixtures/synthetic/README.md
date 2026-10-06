# Synthetic samples
提交自建最小 .docx 与同名说明：生成方法、场景、预期 diagnostic/patch、Word 版本。禁止复制真实论文作为“synthetic”。
现有单元测试运行时创建最小 OOXML ZIP；它只验证加载器，不代替 Word 桌面兼容性验证。第一批可用 Word 样本由 Issue 领取。

Issue #5 修正版的 10 个结构样本位于 [format-samples/](format-samples/README.md)，生成命令从仓库根目录运行，检查已提交文件使用 --check-only。结构校验、CLI 加载与 Word 兼容性分别记录在 [验证记录](format-samples/validation.md)，未检查项不表示通过。
