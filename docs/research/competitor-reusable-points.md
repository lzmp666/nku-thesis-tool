# 竞品可借鉴点（一页）

范围只限 Word 上和排版、样式相关的能力。内容管理、云协作、跨引擎不写。
核实日期 2026-11-10：下表路径逐个查过上游仓库文件树，许可证读的是仓库 `license` 元数据或 LICENSE/COPYING/README 原文，不是转述。
状态：骨架。访谈还没做，这份只收敛「能借鉴什么」，不写产品介绍，也不替访谈下结论。

## 能借鉴什么

| # | 能借鉴什么 | 具体实现位置 | 落到本仓库 | 许可证 | 二创边界 |
|---|---|---|---|---|---|
| 1 | 规则 = 期望格式 + 业务检查 + `enabled` 开关；预警开关按类降噪 | `AfishInLake/WordFormat`：`src/wordformat/rules/node.py`、`docs/rule-handler-guide.md`、`example/undergrad_thesis.yaml` | #111 规则形态 | Apache-2.0 | 借形态，不拷代码 |
| 2 | YAML 点分路径配置（`body.text`、`headings.level_1`）+ 锚点继承 | `src/wordformat/config/models.py` | #111 | Apache-2.0 | 借思路；`general.yaml` 是 draft，不照搬字段 |
| 3 | 题注解析：章号、分隔符、标签与编号间距、续表 | `src/wordformat/rules/caption.py` | #92 | Apache-2.0 | 借解析口径 |
| 4 | 图片宽度按版心（页宽减左右边距）算 | `src/wordformat/rules/object.py` | #92 | Apache-2.0 | 借算法；规范未给 80% 这类数值，不补 |
| 5 | 页面设置 + 页码域 + 首页/奇偶页不同 + 打开时更新域 | `src/wordformat/header_footer.py`、`tests/test_header_footer.py` | #112 已交付（PR #123、#127） | Apache-2.0 | 已实现，只留作对照 |
| 6 | 纯 OMML 段没有 `w:r`，挂在段落遍历下的检查会整段跳过 | `src/wordformat/rules/math.py` | #114 已交付（PR #122、#126） | Apache-2.0 | 这个坑已在 `FormulaAnalyzer` 里绕开 |
| 7 | LaTeX → MathML → OMML 纯 Python 管线 | `src/wordformat/math/omml.py` | 明确不做 | Apache-2.0 | 没有 OMML 回读验证能力，转错没人发现 |
| 8 | 结构候选先按 XML 结构短路，再谈分类；三段式「出结构 JSON → 人工校正 → 校验/修正」 | `src/wordformat/structure/document_builder.py`、`src/wordformat/classify/tag.py`、`docs/领域包开发指南.md` | #111 二期 | Apache-2.0 | 借流程，不引它的分类模型 |
| 9 | 一个目录一个学校、自动发现领域包 | `docs/领域包开发指南.md` | `rules/nku-graduate/colleges/` 目前空着 | Apache-2.0 | 等南开这套规则获批再铺多校 |
| 10 | OOXML 强类型读写 | `dotnet/Open-XML-SDK` | 测试工程已用 3.5.1，引擎仍直接读 XML | MIT | 可直接依赖，`AGENTS.md` 第 8 节已登记 |
| 11 | DOCX 分块先看样式、再看文本、最后回退；metadata 记录样式和页码 | `Unstructured-IO/unstructured`：`unstructured/partition/docx.py` | `structure` 的回退策略参考 | Apache-2.0 | 默认英文词表，中文没评测；只借机制，不引运行时 |
| 12 | DOCX backend 用样式名和 `outlineLvl` 当标题信号 | `docling-project/docling`：`docling/backend/msword_backend.py` | 结构识别对照基线 | MIT（模型权重许可另核） | 只作对照，不替换原文锚点 |
| 13 | 引文样式文件是独立可分发的数据，不是引擎代码 | `citation-style-language/styles`、`citation-style-language/schema` | #119 只查五项一致性，不做 CSL 引擎 | 样式 CC-BY-SA 3.0；schema MIT | 见下节 |
| 14 | 客户端/处理器的许可代价 | `zotero/zotero:COPYING`、`Juris-M/citeproc-js:LICENSE` | 只借「style 机制」这个概念 | Zotero AGPLv3；citeproc-js CPAL-1.0 或 AGPLv3 | 不嵌进闭源 VSTO 插件 |

`AfishInLake/WordFormat` 根目录没有 `NOTICE` 文件，但仍按 Apache-2.0 保留版权与许可证声明、标注修改。上游 tree 共 253 个条目，上面 9 条路径都实际存在。
EndNote 是商业软件，不在开源复用范围，本次没有核它的许可证。

## 许可证结论（能不能直接二创）

| 上游 | 许可证 | 结论 |
|---|---|---|
| `AfishInLake/WordFormat` | Apache-2.0 | 能。可商用、可二创、可闭源分发衍生作品；留版权与许可证声明、标注修改。不传染，不要求本仓库开源 |
| `dotnet/Open-XML-SDK` | MIT | 能。已在测试工程用 |
| `citation-style-language/schema` | MIT | 能 |
| `citation-style-language/styles` | CC-BY-SA 3.0（README 的 Licensing 段） | 有条件。**派生样式必须同许可并署名**，直接把 `.csl` 改一份塞进闭源产物不合适；只读定义当口径参考不涉及派生 |
| `Unstructured-IO/unstructured` | Apache-2.0 | 能，但中文能力要自己评测 |
| `docling-project/docling` | MIT（模型权重另核） | 能，但只在对照位 |
| `zotero/zotero` | AGPLv3（COPYING 原文） | 不嵌。分发衍生作品要按 AGPL 提供对应源码，和闭源插件冲突 |
| `Juris-M/citeproc-js` | CPAL-1.0 或 AGPLv3（LICENSE 原文） | 不嵌。两个选项都带署名/源码提供义务 |

一句话：可以放心二创的是 Apache-2.0 / MIT 那一组；带 ShareAlike 和 AGPL 的只借概念，不进代码和产物。引文这块尤其别把「Zotero 开源」当成「Zotero 代码能搬」——能搬的是 CSL 的样式数据思路，且要看 CC-BY-SA。

## 定位

上游已经把「离线批量改字体行距」做成 `pip install` 就能用的东西。这边再做一遍通用批量排版没有意义。站得住的只有三件：南开 2026 规范的唯一实现、全程走 Word 原生样式加批注、插件内不用切工具。
后面评审新需求先问一句「上游有没有、我们比它强在哪」，答不上来就先不做。

## 待补

- 三类学科真实论文里，哪些开源工具是受访者已经在用的（访谈第 3 问），现在没有数据，不猜。
- 上面第 11、12 条的中文准确率没有在本机样本上测过，`docs/research/unstyled-thesis-open-source-2026-10-26.md` 里也写明上游 benchmark 不能当本插件验收。
