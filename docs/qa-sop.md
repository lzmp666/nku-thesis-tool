# QA SOP
synthetic/：自行构造最小 Word 文档，单一错误到组合错误，包含样式/直接格式、继承、东亚字体、多节边距等；样本生成方法和 Word 版本一并记录。
real-anonymized/：先获得使用许可，再脱敏，再复核；原始文件留在受控存储，不进入 Git。见该目录 README。
goldens/：人工核对的标准记录，不是简单保存引擎输出；采用 template.yaml，稳定记录规则、节点、expected/actual、修复等级与未覆盖项。
回归：restore/build/test；检测正确样本无误报，错误样本有定位；修复后检查内容/关系/样式/节未受损，二次 diagnostic=0、patch=0。suggest/manual 未处理项必须保留，不得假报合规。
需要时比较规范化 XML（命名空间、属性顺序等）并按部件说明合理差异，不用整个 docx 哈希。
Bug：用 Bug 模板，写最小复现、命令、版本、实际/期望、脱敏样本与 golden；截图必须脱敏，不能仅写“不能用”。无法分享真实文档时提供 synthetic 重现。
合并前至少一名非作者 Review；正式规则由负责人审核，未支持项明确记录。

## Bug 模板使用说明
使用 `.github/ISSUE_TEMPLATE/bug.md`，保留背景、工作内容、目录边界、验收和依赖；记录 commit、命令、实际/预期结果、退出码及必要环境信息。优先提供可自行构造的 synthetic 输入。经授权的真实文档按 real-anonymized/ 流程处理，禁止提交原始敏感材料。

## Synthetic 问题试填示例：缺少必需 OOXML 部件

这是当前 CLI 的 synthetic Bug 试填：源码基准版本对空 OOXML 包产生未处理异常，偏离 README 约定的可读错误和退出码 1。无需字体检测插件或真实论文。本 PR 仅完善报告与流程，不修复 Core。

### 背景
验证 CLI 对缺少必需部件的合成 .docx 给出可读错误和非零退出码，不将加载成功等同于格式合规。

### 工作内容
构造一个空 ZIP，扩展名为 .docx，运行当前 CLI，记录输出与退出码并对照基准。若出现成功加载或未处理异常，将本例补齐实际结果后用于 Bug 报告；不为演练修改引擎制造缺陷。

### 版本与环境
- 源码基准：`f22b95ab3940e032f550803fa1132b719bfd4674`（后续复现记录实际 `git rev-parse HEAD`）。
- SDK：.NET 10；执行 `dotnet --version` 记录具体版本。
- 操作系统与 PowerShell：执行 `[Environment]::OSVersion.VersionString` 和 `$PSVersionTable.PSVersion`。
- Word：不适用；本例为加载器负向测试，不验证 Word 桌面兼容性。

### 最小复现步骤与命令
在仓库根目录，用 PowerShell 执行：

```powershell
dotnet restore NkuThesisTool.sln
if ($LASTEXITCODE -ne 0) { throw "restore failed" }
dotnet build NkuThesisTool.sln --no-restore --configuration Release
if ($LASTEXITCODE -ne 0) { throw "build failed" }
dotnet test NkuThesisTool.sln --no-build --configuration Release
if ($LASTEXITCODE -ne 0) { throw "tests failed" }

$samplePath = Join-Path ([IO.Path]::GetTempPath()) ("qa-empty-" + [Guid]::NewGuid().ToString("N") + ".docx")
# 自行构造空 ZIP，不读取或复制任何论文；文件不存在时才创建。
Add-Type -AssemblyName System.IO.Compression
$sampleStream = [IO.File]::Open($samplePath, [IO.FileMode]::CreateNew)
$sampleArchive = [IO.Compression.ZipArchive]::new($sampleStream, [IO.Compression.ZipArchiveMode]::Create)
$sampleArchive.Dispose()
try {
    dotnet run --project src/ThesisEngine.Cli --no-build --configuration Release -- $samplePath
    $sampleExitCode = $LASTEXITCODE
    Write-Output "ExitCode=$sampleExitCode"
} finally {
    # 仅删除本次生成的随机临时文件。
    Remove-Item -LiteralPath $samplePath
}
```

### 实际结果与期望结果 / Golden
基准：stderr 为 `Unable to load document: Missing or duplicate OOXML part: [Content_Types].xml`，`ExitCode=1`，stdout 不出现 `Document loaded successfully`，无未处理异常，不生成修复输出。空 ZIP 没有任何正文、字体或规则，不得报告论文合规。
实际结果（2026-10-04 实测）：从上述 commit 提取原样 Core/CLI 源码与项目文件，在 Windows NT 10.0.26200.0、PowerShell 7.6.5、.NET SDK 10.0.203 下 Release 构建成功（0 warning / 0 error），用相同空 ZIP 输入直接运行构建的 CLI DLL。stderr 为 `Unhandled exception. System.IO.InvalidDataException: Missing or duplicate OOXML part: [Content_Types].xml`，调用链为 DocumentLoader.ReadPart → DocumentLoader.Load → Program，退出码 `-532462766`；未出现加载成功。这是缺陷复现成功，不是修复验收通过。完整 solution 回归以本 PR 的 CI 为准；其他操作系统的异常退出码可能不同，但不得出现未处理异常。重新执行上述完整命令时记录实际 commit、环境、输出和退出码。

### 脱敏声明
输入完全由上述代码生成，不含真实论文、个人信息或涉密数据；本例未附截图。日志分享前检查本机路径并脱敏。

### 风险与回滚
错误接受无效包会让用户误以为文档可继续检测。演练只读加载器，不修改输入、引擎或契约；临时样本在 finally 中删除。未来修复如引入回归，用独立提交 revert，并重新运行回归；不要覆盖原始论文。

### 目录边界
本 Issue #9 仅允许 `.github/ISSUE_TEMPLATE/bug.md`、`docs/qa-sop.md`；生成文件位于系统临时目录，不提交 Git。禁止修改 `contracts/`、`src/ThesisEngine.Core/`、其他 `src/` 及无关目录。发现加载器缺陷另开 Issue 明确范围，由负责人审核；契约表达不足先提 RFC。

### 验收标准
- 模板试填验收：另一位成员按命令自行生成输入、运行 CLI，并能记录和对照明确基准；环境未满足或未执行不能写验收通过。
- 缺陷复现成功：实际结果偏离基准，可在相同 commit/环境重复得到；正常拒绝是回归通过，不是 Bug。
- 若后续需要修复：相同负向输入满足基准，既有合法输入仍可加载，restore/build/test 通过；错误不得变成成功或合规声明。

### 前置依赖
.NET 10 SDK、PowerShell、首次 restore 所需 NuGet 访问，以及当前仓库源码；不依赖 Word 或 VSTO。需要负责人 Review。

## 本次试填验证记录

2026-10-04：按当前 main 的完整 solution、项目、源码及配置重建并执行 Release restore/build/test，原有 5/5 测试通过，0 warning / 0 error；PR #14 仅修改模板和本文，不改变这些执行文件。上述空包异常已复现，真实的 CLI 修复及进程级回归由独立 [PR #15](https://github.com/lzmp666/nku-thesis-tool/pull/15) 交付，新增 9 项，完整 14/14 通过。此记录不表示 PR #15 已合并，也不表示 main 已修复；复现时记录所用 commit。GitHub CI 结果另见各 PR Checks；修复合并后，按同一输入确认可读错误和退出码 1，不得沿用旧异常输出作为新版本基准。

## 真实文件授权 脱敏与复核流程

1. 获取明确授权，并确认许可包含预定使用与仓库提交范围；原件、授权原件及真实身份映射只放受控存储。
2. 在副本上按 [真实样本检查表](../tests/fixtures/real-anonymized/README.md) 处理可见内容、元数据、批注/修订、图片、嵌入对象、所有外部关系与自定义 XML。Word Inspector 后继续逐部件检查；隐藏修订和裁剪图片不等于清除信息。
3. 另存后由第二人复核，记录匿名 sampleId、许可摘要、受控记录编号、执行者/复核者的非敏感标识、Word 版本、日期、未检查项和结果。仓库不得存真实身份映射或敏感授权书。
4. 所有适用项通过且有明确提交许可后，只提交匿名文件和非敏感记录；没有真实样本时模板保持 pending，不填虚构的“已通过”。
