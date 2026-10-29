# 问题检索响应与南开主题更新

本次针对用户反馈的检索点击卡顿和页面不稳定：列表选中立即更新视觉状态，定位请求延后 100ms 并合并连续点击，只对最后选中的问题调用 Word。切换问题仍保留列表顺序与 TopIndex；悬停同步同样不滚动列表。批量修复不受此延时影响。

功能区可用状态在面板状态变化时计算并缓存，回调不再逐个访问 Word 的 ActiveWindow/CustomTaskPanes。只刷新状态确有变化的按钮，不再 Invalidate 整条功能区。页签监听按窗口去重，只在选中页签变化时通知，侧栏可见性相同时不重复设置。显式定位后的短暂窗口内暂停悬停，防止刚选中的问题被鼠标命中的另一个范围替换；鼠标位于检索面板或没有移动时，不进入 Word COM 悬停查询。每次显式定位记录耗时，便于真实长文档复测。

## 设计素材

- 左侧校徽原始文件：`src/ThesisAddIn.VSTO/Assets/nku-emblem-official.jpg`，取自[南开大学官网校徽页](https://www.nankai.edu.cn/164/list.htm)的公开 JPG：[原图](https://www.nankai.edu.cn/_upload/article/images/77/1f/a7852bfa49c3bdea0f4564630e66/c658ec37-5315-4c7c-b5d5-aee3e3e570d3.jpg)。保留原始图案，按 Office 控件尺寸显示。
- 校园主题图：`src/ThesisAddIn.VSTO/Assets/nku-campus-panorama-v2.png`，使用内置 GPT-image 生成，消费于右侧功能区主题图、检索页眉和点击校徽后的主题说明窗口。
- Word 原生功能区的排版与背景由 Office 管理，保留标准命令按钮；项目主题使用青莲紫、暖白和金色，不更改 Office 全局主题。

GPT-image 最终提示词（内置工具模式）：

> Production raster illustration for a Microsoft Word thesis-checking add-in toolbar and search header; wide 3:1 campus illustration, warm ivory #FAF7F1, Nankai lotus purple #762255 and subtle warm gold; fine architectural line art inspired by the Balitai campus historic main building, lotus leaves and open scholarly books; clean editorial illustration, generous negative space, crisp silhouettes, three balanced connected sections; no lettering, numbers, watermark, emblems, fake seals, application mockup or controls.

## 验证

141 项 VSTest 回归通过，签名 VSTO 构建成功。`scripts/Test-WindowsRibbonUi.ps1` 检查状态不变不重绘、局部按钮刷新、getEnabled 回调不再查询状态委托、九个命令图标和两张品牌图片有效；不创建 Word COM 实例。实机点击与安装结果另行记录，不能用该脚本代替 Word 界面验收。

### 本机真实 Word 复验

完整 14 页合成稿，Word 200% Windows 缩放、80% 文档显示：初次分析为 9 项/13 处。第一轮定位到中文字体、西文字体和图表题注的耗时分别为 143、220、189ms；排版更新安装后，中文字体、字号和西文字体三次定位分别为 119、20、76ms（另有 100ms 合并连续点击的等待）。这只是本机该合成稿的测量，不作为所有文档的性能保证。

实测逐项点击时列表顺序和视口保持不变，工具栏关闭/再次打开检索保留选中问题，原生批注连接线与正文高亮仍可见。最后的批量预览显示可修复 13 处、保留 0 项；取消后仍为 9 项/13 处，未写回格式。高 DPI 实测后将数量移到独立浅色条，主题说明窗口按 Word DPI 显式缩放，校徽保留图案和功能区下方“南开大学”文字。

安装在既有 `D:\NkuThesisTool\WordDemo-c2d98261a758` 目录，VSTOInstaller 返回 0，仅注册当前用户。完整截图与预览证据保存在本机 `artifacts/Acceptance-navigation-theme-20261023/`，不将含其他桌面通知的截图提交到公开仓库。该轮仍未独立验收纯鼠标悬停、多文档、长时间使用与全部正式论文规范；PR 和安装包保留草稿/PREVIEW 状态。
