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

`verify_fixtures.py` 分四类判据算，结论分开说：

| 判据 | 覆盖 | 做什么 |
| --- | --- | --- |
| 逐部件重建 | 4 个生成器 / 39 份样本 / 287 个部件 | 导入零依赖生成器在内存里重建样本，比较 ZIP 条目名/顺序与每个部件解压后的字节，不比较 deflate 流、时间戳和外部属性，因此在 Python / zlib 版本不同的机器上不会因压缩流差异误报 |
| 结构断言 | `format-samples/` 10 份样本 | 该目录的重建依赖 python-docx，干净检出无法重建，所以只调用它零依赖的 `verify()` 逐项核对 OOXML 维度。这是**结构断言**，不是逐字节复现 |
| 清单一致性 | 5 个目录 / 51 份 docx | 比对 `manifest.json` 声明、生成器清单、磁盘上的 `.docx` 三个集合，抓「新增样本忘写 manifest」「孤儿 docx」「生成器漏覆盖」 |
| 可选重建 | `format-samples/` 10 份样本 | `--with-rebuild` 额外做一遍逐部件重建校验，需要 python-docx；**默认不开**，缺依赖就跳过并说明 |

`--self-test` 用 9 种情形验证判据既不误报也不漏报，包括「换 deflate 级别和时间戳仍一致」「改一个部件的文字会被发现」「少一个部件会被发现」「manifest 漏声明会被发现」「关闭生成器比对时不误报」「缺依赖时重建校验跳过而非失败」等。

覆盖范围要说清：本脚本**只比对结构**，不替代 Word 桌面版打开副本的兼容性验证，也不等于格式合规结论。`format-samples/` 的重建路径（`generate.py --check`）需要 python-docx==1.2.0（MIT），但**校验已提交样本不需要**：`python3 tests/fixtures/synthetic/format-samples/generate.py --check-only` 只用标准库；`--with-rebuild` 是可选加严档，缺依赖时打印「重建校验已跳过：…需要 docx」并正常退出，不把「没查」写成「查了没问题」。`full-thesis/` 两份 docx 由命令行 `--output` 指定，生成器没有固定清单常量，该目录只比 manifest ↔ 磁盘，不假装也查了生成器。

CI 里若加入，只需在 `build-test` 加一步 `python3 tests/fixtures/synthetic/verify_fixtures.py`。`AGENTS.md` 第 4 节规定 `.github/` 由负责人维护，因此改动 CI 的工作流文件不在这条技术债的 PR 里；逐字节的 `--check` 也不建议直接放进 CI，理由同上。
