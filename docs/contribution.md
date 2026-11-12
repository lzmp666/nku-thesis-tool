# 从领取任务到合并

> 第一次参与，先走一遍[十分钟上手](ai-onboarding.md)：跑通 `analyze --demo` 再回来读这一页。给 AI 干活的话，让它先读根目录 [AGENTS.md](../AGENTS.md)。

1. 接受负责人发来的 GitHub 仓库邀请，打开 Issues，读背景、目录边界、验收与依赖。
2. 在一个无阻塞的 Issue 留言“我领取，预计完成日期…”。负责人确认 assignee 后开工；一人先领取一个任务。
3. 安装 Git 与 .NET 10 SDK（版本以根目录 `global.json` 为准），克隆仓库：
```powershell
git clone https://github.com/lzmp666/nku-thesis-tool.git
cd nku-thesis-tool
git switch main
git pull --ff-only
git switch -c docs/issue-1-rule-notes
```

### 同步基线（远端 main 为唯一基线）

**本地停在旧 commit、远端已前进，就会出现「同一份文件两台机器两个结论」。** 动手前先同步到远端 `main` 最新 commit，并在你的 PR / 核对记录里写明所用 commit SHA：

```powershell
# 查看远端 main 当前指向
git ls-remote https://github.com/lzmp666/nku-thesis-tool.git refs/heads/main

# 已克隆的仓库：硬同步到远端 main（本地 main 不留未推送提交）
git switch main
git fetch origin
git reset --hard origin/main          # 本地 main 与远端对齐；个人分支不受影响
git log -1 --format="%H %s"           # 记下这个 SHA，写进 PR 说明
```

- 2026-11-12 基线：`ed0faf2`（Merge pull request #121，含 #106 的三线表误报修复与 #107 的样式归属检测）。
- 同步后先跑一遍构建与测试确认基线是绿的，再开工。旧 clone 拉不动时可删除后重新克隆。

### 回归基线（三条链路，一条命令）

改完任何格式化逻辑，跑这条命令确认「修干净了」，输出留痕在 `artifacts/regression/`：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Invoke-RegressionBaseline.ps1
```

它依次验证（任何一条不达标都非零退出）：

1. **修复后复检 0 条**：`apply-figure-format` → `analyze-figures`，三线表不再误报「有完整内边框」；
2. **套样式 + 改图表后不误报**：`full-format` → `analyze-figures` / `analyze-styles`，表格 0 条、字体字号无「手动」残留；
3. **二次运行幂等**：对同一输入跑两次 `full-format`，比较解包后的 `word/document.xml`（压掉空白后必须一致）。**不要比整个 docx 的哈希**——zip 条目时间戳每次都不同，按字节比永远过不了。

对应的单元测试在 `tests/ThesisEngine.Tests/RegressionBaselineTests.cs`；合成样本核对见根目录 [AGENTS.md](../AGENTS.md) 第 5 节的 `verify_fixtures.py` 三条命令。
4. 分支前缀只能是 feature/、fix/、rules/、test/、docs/。一个 Issue 一个分支，禁止直接改 main，不建 develop。
5. 只改 Issue 允许目录。发现需要改 contracts/，先提 RFC，等负责人确认；Core 与 contracts PR 必须负责人 Review。
6. 修改后运行 README 的 restore/build/test；规则任务附官方出处、原文、版本、适用范围；样本任务附脱敏证明。
7. 提交并推送：
```powershell
git add <本次修改的文件路径>
git commit -m "docs: add rule notes for issue 1"
git push -u origin docs/issue-1-rule-notes
```
8. GitHub 点击 Compare & pull request，base=main，按模板写，关联 Closes #编号，邀请一人 Review；核心与契约请 @lzmp666。
9. 处理意见、重新自测。至少 1 approval、对话解决、CI 通过后由负责人合并；合并后删短期分支。不要提交真实未脱敏论文、密钥或 bin/obj。

## main 保护（负责人检查）
2026-10-01 初始化检查：GitHub 页面提示当前 Private 仓库套餐不会执行保护，Code Owner review 选项不可用；保存还要求密码确认，因此未保存/启用保护。负责人升级合适套餐并重新认证后手动配置。
在 Settings → Branches → Add classic branch protection rule，匹配 main：
- Require a pull request before merging
- Require approvals：1
- Require review from Code Owners（使 contracts/Core 强制负责人审查）
- Require conversation resolution before merging
- Require status checks：CI / build-test（首次 CI 运行后选择实际检查名）
- Do not allow bypassing the above settings
- Allow force pushes：不选；Allow deletions：不选
CODEOWNERS 只声明归属，必须配合上述规则。私有仓库保护可能需要 GitHub Pro/Team；计划不支持时须升级后开启，期间只能执行团队约定，不能声称已强制保护。
官方说明：https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-protected-branches/about-protected-branches

## 邀请团队
负责人 Settings → Collaborators → Add people，填写成员 GitHub 用户名；成员接受后按步骤 1 领取任务。这里不预填成员、不预分配 Issue。
