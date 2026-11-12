# 项目对话接续笔记

> 最后更新：2026-11-04
> 仓库：https://github.com/lzmp666/nku-thesis-tool
> 本地路径：C:\Users\guo11\Desktop\nku-thesis-tool-main

---

## 一、项目是什么

南开大学研究生论文格式插件。面向南开大学研究生，自动检查/调整论文格式，不修改论文内容。

**第一阶段目标（寒假前交付）：**
1. 基础格式调整：字体字号批量修改 + 撤销
2. 图表排版适配：图片格式统一、表格自动对齐
3. 全文档自动排版：基于 Word 样式体系，一键生成规范文档
4. AI 格式对话：仅回答格式问题，不润色内容

**关键方向：**
- 从"逐段扫描字体"转向"基于 Word 样式体系"（改 styles.xml，全文自动同步）
- 优先后端，前端（VSTO 侧边栏、GUI）后续做
- AI 用免费大模型，Key 从配置读，不硬编码
- 不涉及论文内容修改，只处理形式

---

## 二、当前 open issue（4 个）

| 编号 | 标题 | 状态 |
|---|---|---|
| #90 | Word 样式体系读取与批量重写 | open，待做 |
| #92 | 图表排版适配（图片格式统一、表格自动对齐） | open，待做 |
| #95 | 批量修改的撤销功能 | open，待做 |
| #96 | AI 对话功能（仅格式调整问答） | open，待做 |

### 已关闭的 issue（不要重开）
#87（VSTO 侧边栏）、#88（大模型结构识别）、#89（竞品调研）、#91（GUI 按钮）、#93（安装包）、#94（引文格式）——这些都不是第一阶段范围，已关闭。

---

## 三、已完成的工作

- 字体/字号规则检查（demo/east-asian-font 等）
- Word 批注 + 黄色高亮标注（批注已精简为四项：来源/规则/期望值/实际值）
- CLI 命令行分析工具（`dotnet run -- analyze`）
- Windows Forms GUI（已编译出"南开论文格式工具.exe"）
- VSTO 插件协议层（Protocol 层可运行）
- 2026 版论文规范规则（general.yaml + 覆盖矩阵）
- 安全修复引擎（PatchEngine）
- 所有历史 PR 已合并（#33、#35、#36、#43、#53、#66）

---

## 四、团队成员（6 人）

- Guoooooi（我，项目负责人）
- lzmp666（刘子萌，admin，有合并权）
- magician336
- piggy515
- izuminoseirei
- longzqh319774

---

## 五、技术栈和目录结构

```
src/
  ThesisEngine.Core/       核心引擎（字体分析、批注、补丁引擎、Word读取/写入）
  ThesisEngine.Cli/        命令行工具
  ThesisEngine.App/        Windows Forms GUI
  ThesisAddIn.VSTO/        Word VSTO 插件
  ThesisAddIn.VSTO.Protocol/  插件协议层
  ThesisAddIn.VSTO.ProtocolHost/  协议宿主
tests/ThesisEngine.Tests/  单元测试（当前 63 个）
docs/
  contribution.md          贡献规范
  qa-sop.md               QA 流程
  rule-notes/             规则笔记
```

---

## 六、下次接续从哪里开始

1. 先问我："接着做 #90 吗？"
2. 我会从 #90（Word 样式体系重写）开始写代码
3. 写完后自己 build + test，通过后提 PR
4. 然后依次做 #92、#95、#96

### 注意事项
- git 不在 PATH，用 GitHub Desktop 自带路径：`C:\Users\guo11\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe`
- 取 token：`git credential fill`，拿到后必须 `.Trim()`
- GitHub API 直连不稳定，创建 issue 时要重试
- 提交代码方式：通过 API 建 blob/tree/commit/branch/PR（因为 git push 直连不通）
- 不要在 issue 里提任何人名，用第一人称"我们"
- 严格按照片要求的 4 项功能做，不要自行扩充范围

---

## 七、下次可能要做的决定

- [ ] 从 #90 开始写代码
- [ ] 是否需要我直接开始写后端代码
- [ ] 团队分工怎么安排（群里还没分工）
