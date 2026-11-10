# Synthetic Demo Fixtures

本目录用于 Issue #27 的 Demo Sprint 最小验收，包含：

- `generate_samples.py`：生成原有两份 `.docx`，依赖 `python-docx`
- `generate_checked_sample.py`：零依赖，生成 `clean_checked_sample.docx`，支持 `--check` 逐字节校验
- `clean_sample.docx`：极简样本，段落级 0 条但同时写着「已检查 0 个正文文本片段」，保留作反面教材
- `clean_checked_sample.docx`：真正走完两级检查链路的干净样本
- `known_issues_sample.docx`：含三项已知问题样本
- `manifest.json`：记录每项预期诊断的 `ruleId`、位置、`expected`/`actual`，以及新样本的两级预期

所有样本均为合成演示数据，**仅供演示/测试，不代表合规结论**。规则 ID 以 `demo/` 开头，属于演示规则，不是正式校规；本目录没有修改 `contracts/`、`src/` 或现有学校规则状态。

## 重新生成

原有两份由 `generate_samples.py` 生成，需要 Python 和 `python-docx`：

```bash
python3 -m pip install python-docx
python3 tests/fixtures/synthetic/demo/generate_samples.py
```

该脚本会覆盖当前目录下的 `clean_sample.docx` 和 `known_issues_sample.docx`。本机没装 `python-docx` 也能跑下面这个零依赖脚本，它只碰新样本：

```bash
python3 tests/fixtures/synthetic/demo/generate_checked_sample.py
python3 tests/fixtures/synthetic/demo/generate_checked_sample.py --check
```

`--check` 只比对不写盘，ZIP 条目顺序和时间戳都固定。原有两份当前没有纳入逐字节校验：历史上有过手工调整，重新生成的字节不保证与仓库里的一致，所以改它们之前先看清差异，别拿 `--check` 的结果当判据。

## 两级链路的干净样本

`clean_sample.docx` 段落级的 `诊断数量：0` 是和 `已检查 0 个正文文本片段` 一起出现的——那段文本没有可识别的章节边界，被保留为「未知」，压根没进检查范围；文档级又因为沿用 Word 默认页面设置报 8 处。它适合用来讲「0 条不等于干净」，不适合用来证明检查器工作正常。

`clean_checked_sample.docx` 补的是后一半：

| 层级 | 实跑结果 |
| --- | --- |
| 段落级 `--scope paragraph` | `诊断数量：0`、`已检查 3 个正文文本片段，涉及 3 个确认正文段落`，未覆盖范围只剩「页眉/页脚」 |
| 文档级 `--scope document` | `完整判定 2 / 2 个文字段落`、`已检查的文档级项目均符合。`，无 `⚠️ 不符合规范` |
| 默认 `--scope all` | 两段报告同一份输出里同时成立，exit code 0 |

做法是给正文加可识别的章节边界（第一章 / 参考文献），三段正文显式写宋体 + Times New Roman + 12 磅；页面设置、篇眉和 PAGE 域沿用 `tests/fixtures/synthetic/document/` 里已验证过的规范值。`DemoEndToEndTests` 里锁了覆盖数、完整判定和包结构，防止它退化回「压根没查」。

### Word 16 实机读数

2026-11-10 在 Word `16.0.20527.20150` 里打开该样本的临时副本，`Version=16.0` / `Path=Office16` 确认不是 WPS 的 COM 兼容层（本机 `Word.Application` 的 `LocalServer32` 指向 `wps.exe /Automation`，直接 `New-Object -ComObject` 会拿到 WPS）。

| 项 | 读数 |
| --- | --- |
| 页边距 | 上 3.801 / 下 3.801 / 左 3.200 / 右 3.200 cm |
| 页眉距 / 页脚距 / 装订线 | 3.000 / 3.000 / 0 cm |
| 纸张 | 21.001 × 29.700 cm |
| 正文 | 5 段、1 页，无「无法读取的内容 / 是否修复」提示 |
| 篇眉 | 「第一章 绪论」，宋体 10.5 磅，段落对齐 `1`（居中） |
| 页码 | `Field.Type=33`、`Code=PAGE`、`Result=1`；结果字体宋体 / Times New Roman / 10.5 磅，段落居中 |

整页视图截图见 [docs/demo/word-109-clean-checked.png](../../../../docs/demo/word-109-clean-checked.png)。验证只开临时副本，关闭时不保存，仓库里的样本 SHA-256 前后一致，没有留下 `~$` 锁文件。这些读数证明 Word 能正常解释这份包，不构成规范合规结论。

## 已知问题样本包含

1. 东亚字体问题：预期宋体，实际黑体
2. 西文字体问题：预期 Times New Roman，实际 Arial
3. 字号问题：预期 12pt，实际 10.5pt

`manifest.json` 中的值仅用于演示输入、预期诊断和稳定定位，不代表 2026 版或任何正式论文格式规范。
