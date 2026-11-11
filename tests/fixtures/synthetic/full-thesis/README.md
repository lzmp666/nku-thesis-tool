# 完整合成论文夹具

`synthetic_complete_thesis.docx` 是用于真实 Word 和引擎回归的人工构造样稿，不含学生、导师或真实研究数据。当前 Word 桌面版（16.0.20527）只读打开为 14 页、6,172 字，读数见 `word-readings-2026-11-11.json`；字体、分页和 Word 版本会影响实际页数。

论文内容覆盖中文与英文摘要、目录、六章正文、图、表、公式、参考文献、附录和致谢。正文预设三处错误，另有一条故意错误的图题，用于验证题注规则及批量修复：

- `【中文格式测试点】`：东亚字体设为黑体，预期宋体。
- `【字号格式测试点】`：字号设为 10.5 pt，预期 12 pt。
- `SYNTHETIC ENGLISH FORMAT CHECK`：西文字体设为 Arial，预期 Times New Roman。
- `图4.1　不同时段的合成平均占用水平`：故意设置为黑体 12 pt、左对齐、1.5 倍行距、段前段后 0 pt；预期宋体 10.5 pt、居中、单倍行距、段前 6 pt、段后 12 pt。

分析器应报告三处正文问题及图题的六项可量化格式异常。其他符合规范的图题和表题不应触发诊断。表格单元格字体/边框/表头语义、图片清晰度和图意、题注序号与标题空格及图片/表格的移动仍未自动修复；应保留为明确的人工检查项，不得当作已通过。

## 怎么重建

```bash
python generate.py --check   # 重建两份并与磁盘比整文件字节，不写盘
python generate.py           # 重建并覆盖本目录的两份样本
```

`generate.py` 是本目录的统一入口，形态与其它样本目录一致：`SAMPLES` 清单 + `build_bytes(name)`。它调下面两个生成器：

- `generate_full_thesis.py`：python-docx 造主样本，`build_bytes()` 返回字节，`main()` 写盘。
- `generate_object_layout_errors.py`：从**内存里**的主样本字节派生，`build_bytes(source)` 不读盘；命令行版 `--output` 仍拒绝覆盖原始夹具和已有输出。

依赖 `python-docx`、`Pillow`、`lxml`。`build_bytes()` 的产物会过一遍 `_normalize()`：ZIP 条目顺序原样保留，时间戳统一成 `(2000, 1, 1, 0, 0, 0)`（与 `format-samples/generate.py` 同一个常量），部件字节一字不改。python-docx 的 `save()` 原本取当前时刻当条目时间戳，同一生成器换分钟跑出的整文件字节就不同，哈希没法当判据；固定之后 `--check` 直接比整文件字节。字节不等时脚本会拆到部件级：报出具体哪个部件不同，若部件全同而整文件不同，则说明差异只在容器头部。

整文件字节的可复现范围限于同一台机器：图内文字取自 `C:/Windows/Fonts/msyh.ttc`，换字体、换 Pillow 大版本都会改图片字节。跨机器核对仍按逐部件来（`verify_fixtures.py --with-rebuild` 就是这个口径），别把字体环境差异当样本被改坏。

2026-11-11 实测的版本组合：Python 3.13.14 配 python-docx **1.2.0** 与 **1.1.2**（Pillow 13.0.0、lxml 6.1.3 相同）各自重建，两份样本的整文件字节与仓库里那份完全相同；再把 `format-samples/` 10 份一并比，12 份样本两版本互比也是逐字节一致。组合记在 `manifest.json` 的 `containerNormalization.verifiedByteIdenticalAcrossEnv`。没测过的组合不写「已验证」，只说明按逐部件判据通过。

图内文字用 `C:/Windows/Fonts/msyh.ttc`。换平台或换字体后图片字节会变，重建结果不同属预期，不是样本被改坏。

`generate_unstyled_thesis.py` 的产物没有提交进仓库（只用于手工排查去样式后的表现），所以不在 `SAMPLES` 里。

可用 `scripts/Test-WordInPlaceRepair.ps1` 在 Microsoft Word 中验证全文段落映射、标记、批量修复和保存边界。

## 派生样本

`synthetic_object_layout_errors.docx` 在上述完整合成论文基础上，额外把一幅图片与一张简单固定列宽表格设为 600 pt，故意超出所在分节 456.1 pt 的可用宽度。分析应得到 11 组合并问题。自动化回归验证选择全部补丁后，正文、题注与这两项宽度问题消失，全部文字、表格数量、媒体和其他部件保持原样，原文件字节不变。该测试不代表表格内部字体、边框、图意或图像清晰度已通过检查。

这两个尺寸规则是版面边界检查，不是南开规范要求所有图片或表格具有同一固定宽度。Word 中的实际分页、直接修复与撤销仍须实机验收。

## 2026-11-11 重新产出

仓库里原先那两份样本由更早版本的生成器产出，与当前生成器已经对不上：单元格底纹 `w:shd` 缺 `w:val`、`w:tcPr` 子元素顺序不合 `CT_TcPr` 的 sequence，附录说明段落还是旧稿。`verify_fixtures.py --with-rebuild` 第一次把这两份样本纳入逐部件重建就报了出来，因此按当前生成器重新产出。

改动只有三处（详见 `manifest.json` 的 `regeneration` 字段）：`w:shd` 补 `w:val="clear"`、`w:tcPr` 按 schema 排序、附录那段说明文字更新为当前覆盖范围。段落 184、表格 4、图片 2、节 1、两份媒体字节、14 页页数与全部诊断数都没变。

`docs/demo/` 下 2026-10-23 至 2026-10-26 的实机记录引用的是更早一组字节（`035E0E13…` / `24A430AE…`），那些记录是当时的证据，不改。

同一天还做了第二次重产，只动容器不动内容：`generate.py` 的 `build_bytes()` 加上 `_normalize()` 固定 ZIP 条目时间戳，两份样本的哈希因此从 `8CAE9DEC…` 变成 `90DA9BAA…`，解压后的 20 个部件一字未改（详见 `manifest.json` 的 `containerNormalization`）。当前字节以 `manifest.json` 的 `sha256` 与 `word-readings-2026-11-11.json` 为准。
