# 十分钟上手（人看的，也喂给 AI）

这份文档只干三件事：让你十分钟内跑出第一条检查结论、给你一段能直接粘给 AI 的提示词、把最容易卡住的地方写在前面。

规矩在 [AGENTS.md](../AGENTS.md)，流程细节在 [contribution.md](contribution.md)。这里不重复，只给最短路径。

---

## 1. 跑通这五步

### 先对一下环境

```powershell
dotnet --version
```

输出要是 `10.x`。低于 10，`global.json` 会直接把你挡下来（报 `NETSDK1045`），不用怀疑代码。没装就去 https://dotnet.microsoft.com/download/dotnet/10.0 装 SDK，不要装 Runtime，编译器在 SDK 里。

CLI 在 Windows、macOS、Linux 都能跑。只有两件事要 Windows + 桌面版 Word：打开带批注的标注副本，和装 VSTO 插件。

### 五条命令

```powershell
git clone https://github.com/lzmp666/nku-thesis-tool.git
cd nku-thesis-tool

dotnet restore NkuThesisTool.sln
dotnet build NkuThesisTool.sln --configuration Release --no-restore
dotnet test NkuThesisTool.sln --configuration Release --no-build
```

build 的预期是 `0 个警告 0 个错误`。仓库开了 `TreatWarningsAsErrors`，看到一个警告就是失败，别当成提示。

test 的预期是全过，当前 main 是 **430 条**（2026-11-10）。这个数会随 main 涨；如果你看到明显更少，多半是分支太旧，`git pull` 一下。`AGENTS.md` 第 3 节记着同一条基线，改哪边都要一起改。

然后跑两遍演示，这是第一次真正看到引擎在干活：

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/clean_sample.docx
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx
```

### 第一个应该长这样

下面两块**只是段落级报告的关键行摘录**。实际输出还会打印报告头（`模式`、`输入文件`、`状态`）、`未覆盖范围`、每条问题后面的 `修复级别` 和 `说明`，以及最后一行 `提示`。多出这几行是正常的，逐行比对时只看下面这些字段。

还有一整段不在摘录里：`analyze` 默认 `--scope all`，段落级报告之后会另起一段「文档级版式检查报告」。正确样例是极简合成包，页面设置沿用 Word 默认值（四边 2.54cm、无页眉页脚），所以那段会报 8 处不符合规范。这 8 处和上面的 `诊断数量：0` 不是同一套检查，别当成跑挂了。只跑段落级用 `--scope paragraph`，只跑文档级用 `--scope document`。

```
诊断数量：0
发现问题：0
规则版本：demo-0.1.0
已检查范围：已检查 0 个正文文本片段，涉及 0 个确认正文段落；识别 0 条图/表题注；核对 0 个配对图片/简单表格的宽度
问题详情：
  （无）
```

### 第二个应该长这样

```
诊断数量：3
发现问题：3
问题详情：
  [1] 警告 demo/east-asian-font
      位置：word/document.xml#paragraph[1]/run[0]
      期望：宋体
      实际：黑体
  [2] 警告 demo/western-font
      位置：word/document.xml#paragraph[2]/run[0]
      期望：Times New Roman
      实际：Arial
  [3] 警告 demo/font-size
      位置：word/document.xml#paragraph[3]/run[0]
      期望：12pt
      实际：10.5pt
```

对上了，就算跑通了。

对不上分三种：**只是多出几行**，那是报告头和后面附的文档级报告，不用管；**关键的几个数字不对**（`诊断数量` 不是 0 或 3、`期望/实际` 对不上、`已检查范围` 全是 0 而诊断却是 3），先看第 4 节的环境问题，环境没问题再开 issue 把输出贴上去；**段落级对上了、文档级却报了一串**，看上面那段说明，属于两个口径。


### 一个必须看懂的地方

第一个样本出 `0 条`，同时写着 **「已检查 0 个正文文本片段」**。这两句话要一起读：那个文件里没有可识别的正文起始标题，唯一一段文本被保留为「未知」，压根没进入检查范围。

所以 `0 条` 从来不等于「干净」。这条规矩贯穿整个仓库：**报告里必须能分开「查了，没发现问题」和「没查」**。你以后加检查项也一样，不能只报一个 0。

再往下一层：现在这 `0 条` 只说的是段落级演示规则查到的东西。同一份输出后面还挂着文档级版式报告，那里查的是纸张、页边距、页眉页脚、页码、目录、空白页，和段落级不是一回事。「0 条」既可能是「没查」，也可能只是「这一层没查到」，得看报告写的是哪一层。

### 实际要多久

纯命令时间，2026-11-09 本机从零克隆（不含装 .NET SDK、不含第一次拉 NuGet 包）：

| 步骤 | 用时 |
| --- | --- |
| `git clone`（仓库 7 MB） | 27 秒 |
| `dotnet restore` | 5 秒 |
| `dotnet build` | 8 秒 |
| `dotnet test`（当时 362 条） | 20 秒 |
| 两遍 `analyze --demo` | 3 秒 |
| 合计 | **63 秒** |

这些数字单机单次采样，上下浮动几秒很正常，别拿来当性能基准。表里的测试条数是 2026-11-09 那次实测当时的 main，不是后来的基线，不跟着改。

真正该看的是墙钟：2026-11-10 让一个完全不了解这个仓库的会话，只拿着这份文件从零走了一遍——从敲 `git clone` 到两个 demo 跑完 **95 秒**，算上读文档和踩 clone 证书坑，总共 **2 分 02 秒**。离十分钟还远。

首次装 SDK 和还原 NuGet 包要另算，视网络十几分钟不等。真卡人的不是这几分钟，是第 4 节里那几个报错。


---

## 2. 把这段提示词粘给 AI

换了新会话、或者把活交给 AI 时，整段复制，改两处方括号：

```text
仓库：https://github.com/lzmp666/nku-thesis-tool（.NET 10 解决方案 NkuThesisTool.sln）
任务：GitHub issue #[编号]，标题是 [标题]。
第一步先读仓库根目录的 AGENTS.md，再读 docs/ai-onboarding.md，然后读这个 issue 的正文，尤其是「允许修改目录」和「不允许修改目录」两栏。

开工前必须做：
1. git switch main && git pull --ff-only，然后按 issue 建分支，前缀只能是 feature/ fix/ rules/ test/ docs/。
2. 先跑基线：dotnet restore NkuThesisTool.sln、dotnet build --configuration Release、dotnet test --configuration Release。构建必须是 0 警告 0 错误，测试必须全过。不过就停下来告诉我，不要跳过、不要注释掉测试。
3. 只改 issue 允许的目录。.github/、Directory.Build.props、global.json、NkuThesisTool.sln、NuGet.Config 一律不动；需要动就先问我。

做的过程中：
- 格式问题一律改 Word 样式，不逐段刷直接格式。正文文字、标点、段落顺序、表格内容，以及公式对象里的 m:t，一个字符都不许改。
- 期望值要有出处：哪份文件、哪张表、哪一条。没有出处就不要写进规则，也不要用「一般认为」凑数。
- 检查结论要区分「查了没发现」和「没查」。判不准就报未检查并写清原因，不要写「符合」。
- 样本用对应目录的 generate.py 生成，manifest.json 同步更新，不要直接手改 docx。
- 新增行为要配测试。改完复诊：再跑一遍检查，确认问题消失且没冒出新问题。

交付前：
- 再跑一遍 build + test，贴真实输出（警告数、错误数、通过数）。
- 至少贴一条命令的实跑输出，只写「测试通过」不算。
- PR 按 .github/pull_request_template.md 填，正文写 Closes #[编号]。

边界不清、要改契约、要动的文件不在允许目录里，停下来问我，不要自己决定。
```

---

## 3. 从 issue 到 PR

流程本身在 [contribution.md](contribution.md)，这里是能直接敲的版本。

**领任务**：在 issue 下留言领取，等负责人把你设成 assignee 再开工。一次先领一个。

**开分支**：一个 issue 一个分支，不建 `develop`，不往 `main` 直接推。

```powershell
git switch main
git pull --ff-only
git switch -c docs/issue-109-ai-onboarding
```

**改完自测**：

```powershell
dotnet build NkuThesisTool.sln --configuration Release --no-restore
dotnet test NkuThesisTool.sln --configuration Release --no-build
```

**提 PR**：base 选 `main`，正文按 `.github/pull_request_template.md` 填。有两栏别糊弄：

- 「测试情况」写命令 + 真实输出。build 的警告/错误数、test 的通过数、至少一条命令的实跑结果。
- 「修改范围」写实际改了哪些目录，和 issue 里的允许目录对一遍。

`src/ThesisEngine.Core/` 和 `contracts/` 的改动按 `.github/CODEOWNERS` 要负责人 Review。

### 关于根目录那几个 .ps1

`check_branch.ps1`、`push_branch.ps1`、`push_pr.ps1` 这些是早期直连 GitHub API 的临时脚本，思路可以看——先取 ref 拿 sha，再建分支、推内容。但里面写死了某台机器上的 git 路径和某个具体分支名，换台电脑直接跑会报找不到 `git.exe`。

所以：**照思路，不照抄**。日常用上面的 `git` 命令就够了。哪天要把它改成通用脚本，单独开个 issue，别塞进别的 PR。

---

## 4. 常见问题

### clone 报 `CRYPT_E_NO_REVOCATION_CHECK`

```
fatal: unable to access 'https://github.com/...': schannel: next InitializeSecurityContext failed:
CRYPT_E_NO_REVOCATION_CHECK (0x80092012) - 吊销服务器无法验证证书是否已吊销
```

Windows 上走 schannel 时，网络查不到证书吊销列表就会这样。

**先试这条**，它对单次命令生效，不动全局配置：

```powershell
git -c http.sslVerify=false clone https://github.com/lzmp666/nku-thesis-tool.git
```

能过就完事了。这条的代价是这一次不校验证书，所以只用于 clone 公开仓库这一次，别固化。

如果连它都过不去，才考虑关吊销检查（有些 Windows git 认这个参数，有些版本不认，实测有差异）：

```powershell
git -c http.schannelCheckRevoke=false clone https://github.com/lzmp666/nku-thesis-tool.git
```

能过就固化下来：

```powershell
git config --global http.schannelCheckRevoke false
```

两条都不行，基本是本地根证书或出口被拦，找管理员要证书，不要硬扛。

> 顺序说明：先给的 `sslVerify=false` 是实测在本仓库多次跑通的；`schannelCheckRevoke=false` 更"正规"但换台机器可能不认，所以放后面。踩过一次坑就够了，别两条轮流试。

### 要挂代理

```powershell
git config --global http.proxy http://127.0.0.1:7890
git config --global https.proxy http://127.0.0.1:7890
# 取消
git config --global --unset http.proxy
git config --global --unset https.proxy
```

`dotnet restore` 走的是环境变量，不是 git 的配置，代理要单独设：

```powershell
$env:HTTP_PROXY = "http://127.0.0.1:7890"
$env:HTTPS_PROXY = "http://127.0.0.1:7890"
```

公司网络换了根证书、restore 报证书不受信任时，先试 `$env:NUGET_CERT_REVOCATION_MODE = "offline"`，再考虑往系统证书库里装根证书。

### rebuild 之后行为没变

大概率没重新编译。`dotnet build` 有时认为文件没动过就跳过，改完源码先 `dotnet build -t:Rebuild`，或者删掉对应的 `bin/`、`obj/` 再编。

### 中文输出乱码

CLI 自己把输出编码设成了 UTF-8，但老的控制台代码页不认。先 `chcp 65001` 再跑。还不行就看 `$OutputEncoding` 和终端字体。

### Word 实机怎么验

分两种，别混：

**看标注副本**（CLI 侧，任何 Windows + Word 都行）

```powershell
New-Item -ItemType Directory -Force artifacts | Out-Null
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- analyze --demo tests/fixtures/synthetic/demo/known_issues_sample.docx --annotated-output artifacts/known.annotated.docx
```

然后拿 Word 打开 `artifacts/known.annotated.docx`。该看到三处黄色高亮，每处挂一条中文批注，说明期望值和实际值。输入文件不会被改，已有的输出文件不会被覆盖。

手边没有 Word，也能先确认文件里确实写进去了（不用装任何东西）：

```powershell
$ProgressPreference = 'SilentlyContinue'
Copy-Item artifacts/known.annotated.docx artifacts/known.zip -Force
Expand-Archive artifacts/known.zip artifacts/known -Force
$h = (Select-String -Path artifacts/known/word/document.xml -Pattern 'w:highlight' -AllMatches).Matches.Count
$c = (Select-String -Path artifacts/known/word/comments.xml -Pattern '<comment ' -AllMatches).Matches.Count
"高亮 $h 处，批注 $c 条"
```

期望输出 `高亮 3 处，批注 3 条`。这只证明高亮和批注写进去了，证明不了 Word 里显示得对不对，最终还得用 Word 看一眼。

> `artifacts/` 下的东西不进版本控制：`.gitignore` 里有 `*.docx`，`git status` 看不到刚生成的文件是正常的，`ls artifacts/` 能看到就是生成了。

**验证 VSTO 插件**（任务窗格侧）

按 [src/ThesisAddIn.VSTO/README.md](../src/ThesisAddIn.VSTO/README.md) 做本机签名和当前用户注册，重启 Word，打开一份合成 `.docx`，顶部会出现「论文格式检查」功能区。验收表在 [docs/demo/word-acceptance.md](demo/word-acceptance.md)，任务窗格的实机截图和读数在 [docs/demo/word-109-task-pane.md](demo/word-109-task-pane.md)。

两个本机踩过的坑：

- `Word.Application` 这个 COM 名可能被 WPS 注册走。先看 `$w.Path`，如果指向 `wps.exe` 所在目录，那 `New-Object -ComObject Word.Application` 拿到的是 WPS 而不是 Word（`Version=14.0`、`Name` 仍显示 `Microsoft Word`）。这时候别拿 COM 的 `PageSetup`、`Field` 读数去写 Word 验收记录，直接启动 `Office16\WINWORD.EXE` 打开文档，按窗口标题区分。
- 插件直接改**当前打开的那份文档**。要验就先把合成样本复制到临时目录再打开，别拿仓库夹具直接跑「分析当前文档」；关掉时选不保存，收尾核对原文件哈希。

这是两套不同的写入边界：CLI 生成新文件，插件直接改当前打开的文档、改动可撤销、保存由作者自己决定。写验收记录时别把两边的结论混着写。

### `undo` 撤销不了

撤销记录存在系统临时目录下的 `nku-thesis-tool-backups`（不在仓库里）。清过临时目录、或者换了机器，记录就没了。

`undo` 认的是**输出文件**，不是源文件：

```powershell
dotnet run --project src/ThesisEngine.Cli/ThesisEngine.Cli.csproj --configuration Release --no-build -- undo path/to/output_nku.docx
```

不确定有哪些，先 `undo --list` 看全部记录（会列出改动说明、源文件、输出文件）。

### 改了 docx 样本，测试还是找不到文件

`.gitignore` 会忽略 `*.docx`。合成样本是逐个目录加白名单放行的。新建样本目录时，除了 `generate.py` 和 `manifest.json`，还要在 `.gitignore` 里加一行 `!tests/fixtures/synthetic/<目录>/*.docx`，并且只 `git add` 具体文件，别图省事 `git add -A`。加完 `git status` 里必须能看到那些 docx。

---

## 5. 这份文档的维护

实测环境：Windows 11 + .NET SDK 10.0.401，2026-11-09；2026-11-10 按 main 复核了命令、430 条基线和 `--scope all` 的实际输出，补了 Word 16 任务窗格实机截图。

命令、测试数、报错原文都会过期。改 `AGENTS.md` 或这份文档时，命令必须当场跑一遍再写进去，别凭印象改参数。发现文档和实际输出对不上，直接开 issue——文档写错比没写更费时间。
