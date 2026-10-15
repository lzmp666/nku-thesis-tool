# 南开大学研究生学位论文格式整改工具
Private repository: https://github.com/lzmp666/nku-thesis-tool

## 目标与当前阶段
Phase 0：可运行项目骨架、契约、协作流程已初始化。Word (.docx) 为主；独立核心引擎，CLI 验证 MVP，VSTO 作为后续插件壳。
当前 CLI 仅验证基本 OOXML 包结构，不检测格式、不修复、不声称论文合规。学校正式规范尚待逐条核验，规则库保持 draft。

## 架构
Official Specification → Rule YAML → Document Reader → Diagnostic Engine → Patch Engine → Report → CLI/VSTO。
修复闭环：诊断 → OOXML 补丁 → 再诊断 → 合规报告。三级修复 auto / suggest / manual。
v1：字体、字号、行距、页边距。后续：标题层级/编号、目录/页码、题注/交叉引用、参考文献。
契约唯一权威位于 contracts/，变更走轻量 RFC。语义幂等要求第二次 diagnostic=0、patch=0；必要时比较规范化 OOXML，禁止使用整个 docx 哈希判定。

## 开发环境与运行
安装 .NET 10 SDK。Core 与 CLI 使用 net10.0；VSTO 不纳入此解决方案，参见 src/ThesisAddIn.VSTO/README.md。
```powershell
dotnet restore NkuThesisTool.sln
dotnet build NkuThesisTool.sln --no-restore
dotnet test NkuThesisTool.sln --no-build
dotnet run --project src/ThesisEngine.Cli -- "C:\path\thesis.docx"
```
成功输出 Document loaded successfully，随后显示“占位模式，尚未检测格式”。缺少参数返回 2，文件或 OOXML 包无效返回 1。
首次 restore 需要连接 NuGet。单元测试自动生成临时 synthetic 包，不使用真实论文。

演示规则分析和标注副本：

```powershell
dotnet run --project src/ThesisEngine.Cli -- analyze --demo "C:\path\known_issues_sample.docx" --annotated-output "C:\path\known_issues_sample.annotated.docx"
```

命令只读取输入文件，将诊断结果输出到终端；传入 `--annotated-output` 时另存带黄色高亮和中文 Word 批注的副本，不覆盖输入文件。该模式使用 `demo-0.1.0` 草稿规则，仅供演示和测试，不代表合规结论。

执行 Demo 自动修复（编号从 1 开始，或使用 `--all` 批量执行）：

```powershell
dotnet run --project src/ThesisEngine.Cli -- repair --demo "C:\path\known_issues_sample.docx" --output "C:\path\known_issues_sample.fixed.docx" --all
dotnet run --project src/ThesisEngine.Cli -- repair --demo "C:\path\known_issues_sample.docx" --output "C:\path\known_issues_sample.fixed.docx" --patch 1
```

修复始终另存新文件；批量模式只执行 `auto` 补丁，`suggest`/`manual` 诊断会在修复后报告中保留。

## 目录与贡献
contracts/ 契约；rules/ 学校规则；src/ 引擎与入口；tests/ 样本、golden、测试；docs/ 协作说明。
只有 main 是长期分支。领取 Issue → 建分支 → 自测 → PR → Review → Merge。禁止直接推 main。
先阅读 [贡献指南](docs/contribution.md)、[规则指南](docs/rule-guide.md)、[QA SOP](docs/qa-sop.md)。
首批 Issue 暂不分配，成员在 Issue 留言领取，由负责人确认 assignee，避免重复工作。

## 初始化待办
- main 保护是否生效见 docs/contribution.md；CODEOWNERS 文件本身不会阻止合并。
- 负责人邀请 4 名成员，确认每人能访问仓库并领取一个无前置依赖任务。
- 规则负责人核验官方来源、版本和适用范围后，才可批准正式规则。
- 暂不选许可证；真实论文必须先取得授权、脱敏并复核，禁止提交原始敏感材料。

