# Synthetic samples
提交自建最小 .docx 与同名说明：生成方法、场景、预期 diagnostic/patch、Word 版本。禁止复制真实论文作为“synthetic”。
现有单元测试运行时创建最小 OOXML ZIP；它只验证加载器，不代替 Word 桌面兼容性验证。第一批可用 Word 样本由 Issue 领取。

Issue #5 修正版的 10 个结构样本位于 [format-samples/](format-samples/README.md)，生成命令从仓库根目录运行，检查已提交文件使用 --check-only。结构校验、CLI 加载与 Word 兼容性分别记录在 [验证记录](format-samples/validation.md)，未检查项不表示通过。

## 样本生成与校验

`document/`、`formula/`、`demo/` 的生成器只用 Python 标准库，各自的 `--check` 逐字节核对本目录样本，适合本机复现。跨机器核对用下面这条，从仓库根目录跑：

```bash
python3 tests/fixtures/synthetic/verify_fixtures.py
python3 tests/fixtures/synthetic/verify_fixtures.py --self-test
python3 tests/fixtures/synthetic/verify_fixtures.py --with-rebuild
```

`verify_fixtures.py` 分四类判据算，结论分开说：

| 判据 | 覆盖 | 做什么 |
| --- | --- | --- |
| 逐部件重建 | 4 个零依赖生成器 / 39 份样本 / 287 个部件 | 导入生成器在内存里重建样本，比较 ZIP 条目名/顺序与每个部件解压后的字节，不比较 deflate 流、时间戳和外部属性，因此在 Python / zlib 版本不同的机器上不会因压缩流差异误报 |
| 结构断言 | `format-samples/` 10 份样本 | 该目录的重建依赖 python-docx，干净检出无法重建，所以只调用它零依赖的 `verify()` 逐项核对 OOXML 维度。这是**结构断言**，不是逐字节复现 |
| 清单一致性 | 5 个目录 / 51 份 docx | 比对 `manifest.json` 声明、生成器清单、磁盘上的 `.docx` 三个集合，抓「新增样本忘写 manifest」「孤儿 docx」「生成器漏覆盖」 |
| 可选重建 | `format-samples/` 10 份 + `full-thesis/` 2 份 | `--with-rebuild` 额外做一遍逐部件重建校验，前者需 python-docx，后者还需 Pillow 与 lxml；**默认不开**，缺依赖就整组跳过并说明 |

`--self-test` 用 12 种情形验证判据既不误报也不漏报，包括「换 deflate 级别和时间戳仍一致」「改一个部件的文字会被发现」「少一个部件会被发现」「manifest 漏声明会被发现」「关闭生成器比对时不误报」「缺依赖时重建校验跳过而非失败」「样本名不带 `.docx` 时补齐后缀」「未知样本名会报错而不是回落」「依赖缺一个时该组整组跳过」。

覆盖范围要说清：本脚本**只比对结构**，不替代 Word 桌面版打开副本的兼容性验证，也不等于格式合规结论。`format-samples/` 的重建路径（`generate.py --check`）需要 python-docx==1.2.0（MIT），但**校验已提交样本不需要**：`python3 tests/fixtures/synthetic/format-samples/generate.py --check-only` 只用标准库；`--with-rebuild` 是可选加严档，缺依赖时打印「重建校验已跳过：…」并正常退出，不把「没查」写成「查了没问题」。

六个生成器同形：清单常量叫 `SAMPLES`，构建函数叫 `build_bytes(name)`（`formula/generate.py` 里叫 `build(name)`），都吃样本名、都返回 docx 字节。`sample_file()` 统一补 `.docx` 后缀——`document/`、`formula/`、`full-thesis/` 的清单写全名，`format-samples/` 只写主题名，路径拼接不再一处带一处不带。

`full-thesis/` 的 `generate.py` 顶层只用标准库（python-docx、Pillow、lxml 都延迟到 `build_bytes()` 里导入），所以清单核对在干净检出下也能跑；重建那一步才需要依赖。

CI 里若加入，只需在 `build-test` 加一步 `python3 tests/fixtures/synthetic/verify_fixtures.py`。`AGENTS.md` 第 4 节规定 `.github/` 由负责人维护，因此改动 CI 的工作流文件不在这条技术债的 PR 里；逐字节的 `--check` 也不建议直接放进 CI，理由同上。CI 若也要跑 `--with-rebuild`，得先装 python-docx、Pillow、lxml。

Issue #111 的规则分析回归样本 `analysis_compliant.docx` / `analysis_noncompliant.docx` 由 [generate_analysis_samples.py](generate_analysis_samples.py) 生成：前者预期 0 条不符合且未检查非空，后者第一段故意偏离字体、字号、对齐、缩进、行距和段前后，预期 8 条不符合。样本均为合成数据。
