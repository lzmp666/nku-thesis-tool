# Word 16 任务窗格实机截图与读数（Issue #109）

验证日期：2026-11-10。补的是 #109 收尾时留下的 `docs/demo/` 截图，拍的是 VSTO 任务窗格，不是 CLI 标注副本。

## 环境

| 项 | 读数 |
| --- | --- |
| Word | `Office16\WINWORD.EXE`，`16.0.20527.20150` |
| VSTO 加载项 | `ThesisAddIn.VSTO` `1.5.4.0`，当前用户 LocalAppData 下的 `NkuThesisTool` 安装目录，注册表 `LoadBehavior=3` |
| 输入 | `tests/fixtures/synthetic/full-thesis/synthetic_complete_thesis.docx` 的临时副本，14 页 / 6154 字 |
| 输入 SHA-256 | `035e0e13780a0269c793e45bb059c1764761aa22d7326e871539abef6a16e4f0` |
| 规则 | 演示规则，不构成南开规范合规结论 |

### 一个环境坑：`Word.Application` COM 被 WPS 接管

本机 `HKCR\Word.Application\CLSID` 的 `LocalServer32` 指向 WPS 的 `wps.exe /Automation`。所以 `New-Object -ComObject Word.Application` 拿到的是 WPS，不是 Word：`Version=14.0`、`Name=Microsoft Word`、`Path=...\office6`，窗口标题带 `WPS Office`。只看 `Name` 会以为拿到了 Word。

这次没有用 COM 做 Word 实机验收，改为直接启动 `Office16\WINWORD.EXE` 打开文档，用窗口类名 `OpusApp` 加标题里的 ` - Word` 区分两者。**以后凡是靠 COM 读 `PageSetup`、`Field`、`Content` 的验收记录，先查 `Application.Path`**，否则读数来源可能记错。

## 实机读数

插件日志原文（`addin.log`，同一轮）：

```
实际加载插件：version=1.5.4.0
插件启动完成；引擎已配置：是；问题检索列表：已创建
收到“分析当前文档”点击。
已为 Word 窗口 460804 生成当前内容快照，开始分析。
已在当前 Word 文档刷新标注，问题数 9，窗口 460804；未打开或保存其他文档。
问题定位耗时：52ms；位置=word/document.xml#paragraph[32]/run[1]
```

界面读数：

| 位置 | 读数 |
| --- | --- |
| 功能区 | 「论文格式检查」选项卡出现；分析当前文档、关闭问题检索、解决当前问题、批量修复、上一项/下一项等按钮 |
| 左侧「问题检索」 | 分析后「待处理 9 项 + 已修复 0 项」；卡片可读到要求/实际值，例如中文字体 `宋体 / 黑体`、字号 `12pt / 10.5pt`、西文字体 `Times New Roman / Arial` |
| 右侧批注窗格 | 9 项，作者 `NKU Thesis Tool` |
| 点第一张卡片后 | 卡片进入选中态；状态栏「第 5 页，共 14 页」「已定位到第 33 段 第 2 个文本片段：中文字体」；修复组按钮转为可用 |
| 缩放 | 全文视图 10%，定位视图 80% |

问题数 9 是演示规则下这 14 页合成稿的诊断条数，不是规范条款覆盖率。

## 两张截图

- [`word-109-task-pane.png`](word-109-task-pane.png)：80% 视图。左侧 9 项问题卡片、第一张选中态、正文第 1 章绪论里的定位位置、右侧 9 条批注、功能区修复按钮可用。
- [`word-109-task-pane-overview.png`](word-109-task-pane-overview.png)：10% 全文视图。14 页缩略排版、左侧「待处理 9 项」、右侧批注「9 项」同屏。

两张都是 1440 × 900 屏幕截图，没有裁剪或合成，没有桌面通知、真实论文或个人信息。正文是合成测试稿，封面写明「本文作为软件测试用虚构论文，不含真实学生研究或真实调查数据」。

### 截图里一处如实保留的现象

右侧批注窗格每条批注下方显示 Word 原生的「缺少内容 / 此注释线包含尚不支持的内容，选择『已在修订』窗格中查看它」。本次**没有**把它解释成插件故障，也**没有**声称已解决，先记为待复核：左侧问题卡片能读到要求/实际值，卡片定位与修复按钮都可用，但批注窗格的这段渲染提示需要在后续界面验收里单独看。不把它从截图里裁掉。

## 写入边界

- 只打开临时副本。「分析当前文档」写入的是内存中副本的标注，关闭时选「不保存」。
- 前后核对仓库夹具 SHA-256 同为 `035e0e13…`；`tests/fixtures/synthetic/demo/known_issues_sample.docx` 也核对过，均未被改写。
- Word 与 WPS 产生的 `~$` 锁文件已清理，仓库工作区没有多出来的 docx。
- 没有调用任何保存命令，没有改插件安装状态或注册表。
