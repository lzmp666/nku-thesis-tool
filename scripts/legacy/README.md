# 历史调试脚本（仅供参考，不要执行）

这个目录收的是历史上排 CI、推分支时留下的一次性调试脚本。它们大多写死了
某台机器的 git 路径、某个具体分支（如 `feature/issue-96-chat`）、某次 CI
run 的 id 或某个人的 GitHub 用户名，换个环境必然跑不通，也不打算再修。

需要查 CI 状态、建远端分支或走 API 开 PR 时，用上一级的正式工具：

| 正式工具 | 用途 |
| --- | --- |
| `../check_branch.ps1` | 只读查询某个分支 ref 的 sha |
| `../push_branch.ps1` | 从基线分支幂等创建远端分支（`-DryRun` 只做 GET） |
| `../push_pr.ps1` | git push 被封时走 GitHub 数据 API 上传文件并开 PR |
| `../Invoke-GitHubApi.ps1` | 上述脚本共用的 API 封装（凭据、repo/branch 解析） |

本目录里的文件仅作为排障记录保留，随时可能整目录删除。
