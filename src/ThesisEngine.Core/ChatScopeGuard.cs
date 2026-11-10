namespace ThesisEngine.Core;

/// <summary>
/// 范围硬约束：用代码强制聊天只回答论文格式问题，不写内容、不润色、不翻译、不回答无关话题。
/// 这是“代码级强制检查”，而不是只依赖提示词——即使用户用话术绕开，代码也会拒绝。
/// </summary>
public static class ChatScopeGuard
{
    /// <summary>
    /// 若问题在格式范围内返回 null；否则返回明确的拒绝理由（说明本助手只回答哪些格式主题）。
    /// </summary>
    public static string? RejectIfOutOfScope(string question)
    {
        if (string.IsNullOrWhiteSpace(question))
            return "请输入你想咨询的论文格式问题。";

        if (FormatKnowledge.TryAnswer(question) is not null)
            return null;

        return "我只能回答南开大学研究生学位论文的【格式】问题（字体、字号、对齐、行距、标题、图表、摘要、参考文献等），" +
               "不撰写、不润色、不翻译论文内容，也不回答与格式无关的话题。请换一个格式相关问题。";
    }
}
