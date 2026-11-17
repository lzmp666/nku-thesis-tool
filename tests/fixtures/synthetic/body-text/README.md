# R11 正文段落格式检查合成样本

对应 `general.yaml` R11（4.7 表4.6，来源 S2026 2026 版，#43 已核对，规则状态 draft）。

`analyze-body-text` 只补齐 #43 尚未覆盖的四个排版维度：**两端对齐、首行缩进 2 汉字符、固定值行距 20 磅、段前后 0 磅**。字体与字号（宋体 / Times New Roman 12 磅）由 #43 的 `NkuGraduate2026BodyAnalyzer` 负责，本批样本全部用合规的宋体 / Times New Roman 12 磅，不制造该维度的差异，避免两套检查重复报同一问题。

## 样本清单

| 文件 | 预期 |
| --- | --- |
| `body_compliant.docx` | 0 问题；正文边界识别为章标题第 1 段 → 参考文献第 5 段之前，3 段正文全合规 |
| `body_align_left.docx` | 3×T1（正文未两端对齐） |
| `body_indent_wrong.docx` | 3×T2（首行缩进不是 2 汉字符） |
| `body_line_spacing_wrong.docx` | 3×T3（行距非固定 20 磅） |
| `body_before_wrong.docx` | 3×T4（段前非 0） |
| `body_after_wrong.docx` | 3×T4（段后非 0） |
| `body_math_line_exempt.docx` | 0 问题；含数学表达式段的行距按条文豁免，列未判定；对齐/缩进/段距仍合规 |
| `body_no_boundary.docx` | 0 问题；没有章标题，正文起点无法确定，如实列未检查（checked=0），不假装合规 |

## 口径边界

- 正文区间：第一个章标题之后、参考文献/附录/致谢等结束标题之前。边界识别不到就列未检查，不拿封面/摘要冒充正文。
- 标题段落、图表题注、表格内段落不是正文，跳过（各有专属规则）。
- 含数学表达式的段落只豁免行距，对齐/缩进/段距照查。
- 只认样式链能展开到的显式值；判不出列未判定，不拿默认值凑问题数。

零依赖生成器 `generate.py`（仅标准库），ZIP 条目顺序与时间戳固定；`python generate.py --check` 逐字节校验。已纳入 `verify_fixtures.py` 的逐部件重建与清单一致性核对。样本均为合成数据，规则状态 draft，不代表 Word 实机渲染正确。
