# Synthetic Demo Fixtures

本目录用于 Issue #27 的 Demo Sprint 最小验收，包含：

- `generate_samples.py`：可重复生成两份 `.docx`
- `clean_sample.docx`：无格式问题样本
- `known_issues_sample.docx`：含三项已知问题样本
- `manifest.json`：记录每项预期诊断的 `ruleId`、位置、`expected`/`actual`

所有样本均为合成演示数据，**仅供演示/测试，不代表合规结论**。规则 ID 以 `demo/` 开头，属于演示规则，不是正式校规；本目录没有修改 `contracts/`、`src/` 或现有学校规则状态。

## 重新生成

需要 Python 和 `python-docx`：

```bash
python3 -m pip install python-docx
python3 tests/fixtures/synthetic/demo/generate_samples.py
```

生成脚本会覆盖当前目录下的 `clean_sample.docx` 和 `known_issues_sample.docx`。

## 已知问题样本包含

1. 东亚字体问题：预期宋体，实际黑体
2. 西文字体问题：预期 Times New Roman，实际 Arial
3. 字号问题：预期 12pt，实际 10.5pt

`manifest.json` 中的值仅用于演示输入、预期诊断和稳定定位，不代表 2026 版或任何正式论文格式规范。
