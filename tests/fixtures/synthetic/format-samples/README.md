# 合成格式样本

Issue #5 的 10 个自建结构样本，未复制真实论文。基线是假设测试值，不是南开校规；正式规则与学校来源审核待完成。当前引擎不检测格式，以下是人工预期，不是引擎运行结果。

## 假设基线与生成

中文宋体、西文 Times New Roman、正文 12pt、行距 1.5 倍；页边距上/下/右 2.5cm、左 3cm。标题黑体 16pt、居中、不加粗。规范修改时同时更新 BASE、构造器/断言及本文，不声称仅改一个字典即可完成所有更新。

从仓库根目录执行（Python 3.9+）：

```sh
python -m pip install -r tests/fixtures/synthetic/format-samples/requirements.txt
python tests/fixtures/synthetic/format-samples/generate.py --check
python tests/fixtures/synthetic/format-samples/generate.py --check-only
```

--check 会重新生成；--check-only 只读取已提交样本。ZIP 排序、时间戳和非个人元数据固定；相同 Python/python-docx 环境下重复生成可逐字节核对。校验所有样本的字体、字号、行距、全部节边距、文本、对照继承和唯一覆盖属性，不以生成成功等于验证成功。

## 场景与定位

每个错误文件只偏离一个格式维度，其余属性按基线核对。正文按“正文段”计数，标题不计入；在 document.xml 中首标题为 p[1]，正文第二段为 p[3]。

| 文件名（均为 .docx） | 输入偏差 | 人工预期 |
| --- | --- | --- |
| baseline_correct | 无 | 仅本文列出的基线维度匹配，不表示全面合规 |
| font_family_cn_wrong | 所有正文中文楷体 | 定位正文 run 中文字体不匹配 |
| font_ascii_wrong | 所有正文西文 Arial | 西文/数字字体不匹配；不误报中文字体 |
| font_size_wrong | 所有正文 10.5pt | 直接 sz=21，目标 sz=24；字号不匹配 |
| line_spacing_wrong | 正文单倍 | 正文行距不匹配，标题仍 1.5 倍 |
| margin_wrong | 第1节左 2cm | 第1节左边距不匹配，其他边距不变 |
| style_override_control | 正文 run 无 rPr | 从 Normal 继承宋体/Times New Roman/12pt；改变 Normal 后正文随之改变 |
| style_override_wrong | 仅正文第二段 run 写 eastAsia=楷体 | document.xml 的 p[3]/r[1] 中文字体不匹配；其他 run 仍继承；改变 Normal 中文字体时该段维持楷体 |
| multisection_control | 两节均正确 | 两节全部边距匹配；必须遍历全部节 |
| multisection_margin_wrong | 仅第2节左 2.5cm | 只报告第2节左边距，不能误报第1节 |

结构样本未绑定运行 ruleId，诊断和补丁需要实现与独立审核。确定错误建议 suggest，用户未采纳时无可执行 patch；未知/不支持项记录 manual，不因零诊断宣称合规。字体安装、字体替代和 Word 布局不由 XML 名称校验证明。

## 手工重建与 Word 复核

在 Word 新建文档，按基线设置 Normal 中西文字体与字号，设置段落行距和四边页边距，输入 generate.py 的 TEXT 及标题。错误文件只改变表中一项；继承对照不添加 run 直接格式；覆盖错误只对第二正文段中文字体设楷体；多节文件插入下一页分节符后分别设置两节边距。另存 docx。

Word 验证记录见 validation.md。目标环境为 Word 2016+ Windows，目标兼容性不等于实测结论。验证时记录具体版本/build、系统、字体是否安装及每个文件结果；打开无修复提示后检查字体/段落/页面设置和继承行为。未实测范围必须保留 pending。
