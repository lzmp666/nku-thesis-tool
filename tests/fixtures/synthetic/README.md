# Synthetic samples
提交自建最小 .docx 与同名说明：生成方法、场景、预期 diagnostic/patch、Word 版本。禁止复制真实论文作为“synthetic”。
现有单元测试运行时创建最小 OOXML ZIP；它只验证加载器，不代替 Word 桌面兼容性验证。第一批可用 Word 样本由 Issue 领取。

Issue #5 修正版的 10 个结构样本位于 [format-samples/](format-samples/README.md)，生成命令从仓库根目录运行，检查已提交文件使用 --check-only。结构校验、CLI 加载与 Word 兼容性分别记录在 [验证记录](format-samples/validation.md)，未检查项不表示通过。

## 样本生成与校验

`document/`、`formula/`、`demo/` 的生成器只用 Python 标准库，各自的 `--check` 逐字节核对本目录样本，适合本机复现。跨机器核对用下面这条，从仓库根目录跑：

```bash
python3 tests/fixtures/synthetic/verify_fixtures.py
python3 tests/fixtures/synthetic/verify_fixtures.py --self-test
```

`verify_fixtures.py` 导入四个生成器在内存里重建样本，比较 ZIP 条目名/顺序与每个部件解压后的字节，不比较 deflate 流、时间戳和外部属性，因此在 Python / zlib 版本不同的机器上不会因压缩流差异误报。当前覆盖 4 个生成器、39 份样本、287 个部件。`--self-test` 用「换压缩级别和时间戳」「改掉一个部件的文字」「少一个部件」三种情形验证判据既不误报也不漏报。

覆盖范围要说清：`format-samples/` 那套依赖 `python-docx` 的生成器**不在**本脚本覆盖内，通过这条命令不等于全仓库样本都验过；那份样本仍按它自己的 `--check-only` 流程走。另外，本脚本只比对结构，不替代 Word 桌面版打开副本的兼容性验证。

CI 里若加入，只需在 `build-test` 加一步 `python3 tests/fixtures/synthetic/verify_fixtures.py`。`AGENTS.md` 第 4 节规定 `.github/` 由负责人维护，因此改动 CI 的工作流文件不在这条技术债的 PR 里；逐字节的 `--check` 也不建议直接放进 CI，理由同上。
