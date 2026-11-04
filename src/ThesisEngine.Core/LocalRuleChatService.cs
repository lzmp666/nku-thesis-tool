namespace ThesisEngine.Core;

/// <summary>
/// 本地规则版聊天服务。
/// 第一阶段先做基础规则匹配，后续替换为外接大模型。
/// </summary>
public sealed class LocalRuleChatService : IChatService
{
    public Task<string> AskAsync(string question, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(question))
        {
            return Task.FromResult("请输入你想问的格式问题。");
        }

        var q = question.Trim();

        // 简单关键词匹配，后续替换为大模型
        if (q.Contains("一级标题") || q.Contains("Heading 1") || q.Contains("1级标题"))
        {
            return Task.FromResult("一级标题应该用黑体三号，居中对齐。");
        }
        if (q.Contains("二级标题") || q.Contains("Heading 2") || q.Contains("2级标题"))
        {
            return Task.FromResult("二级标题应该用黑体四号。");
        }
        if (q.Contains("正文") || q.Contains("宋体"))
        {
            return Task.FromResult("正文中文用宋体小四，英文用 Times New Roman 小四，1.5 倍行距。");
        }
        if (q.Contains("行距") || q.Contains("行间距"))
        {
            return Task.FromResult("正文行距应为 1.5 倍行距。");
        }
        if (q.Contains("图表") || q.Contains("图片") || q.Contains("表格"))
        {
            return Task.FromResult("图片应居中，宽度不超过正文宽度的 80%；表格应为三线表，表题在上方居中。");
        }
        if (q.Contains("参考文献") || q.Contains("引用"))
        {
            return Task.FromResult("参考文献格式应符合 GB/T 7714 规范。");
        }

        return Task.FromResult("我目前只能回答格式相关问题。你可以问我：标题用几号字？正文用什么字体？行距多少？图表怎么排版？");
    }
}
