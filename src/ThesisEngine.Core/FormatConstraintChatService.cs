namespace ThesisEngine.Core;

/// <summary>
/// 离线规则版聊天服务（无需任何 API Key）。在代码中强制范围硬约束与引用约束，
/// 所有格式回答都来自 <see cref="FormatKnowledge"/> 的权威条目（含对齐等硬约束）。
/// 作为 <c>--local</c> 离线模式使用。
/// </summary>
public sealed class FormatConstraintChatService : IChatService
{
    public Task<string> AskAsync(string question, CancellationToken cancellationToken = default)
    {
        var rejection = ChatScopeGuard.RejectIfOutOfScope(question);
        if (rejection is not null)
            return Task.FromResult(rejection);

        var q = (question ?? string.Empty).Trim();
        if (CitationConstraint.IsCitationQuestion(q))
        {
            var extracted = TryExtractCitation(q);
            if (extracted is not null)
            {
                var problem = CitationConstraint.ValidateCitation(extracted);
                if (problem is not null)
                    return Task.FromResult("引用不符合 GB/T 7714 硬约束：" + problem + " 请按规范修正后再提交。");
            }

            return Task.FromResult(CitationConstraint.CanonicalAnswer);
        }

        var answer = FormatKnowledge.TryAnswer(q);
        return Task.FromResult(answer ?? "我目前只能回答格式相关问题。你可以问我：标题用几号字？正文用什么字体？行距多少？图表怎么排版？");
    }

    private static string? TryExtractCitation(string question)
    {
        var start = question.IndexOf('[');
        if (start < 0)
            return null;

        var end = question.IndexOfAny(new[] { '?', '？', '\n', '\r' }, start);
        var snippet = end < 0 ? question[start..] : question[start..end];
        return snippet.Length > 4 ? snippet : null;
    }
}
