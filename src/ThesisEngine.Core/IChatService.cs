namespace ThesisEngine.Core;

/// <summary>
/// 论文格式对话服务接口。
/// 只回答格式调整相关问题，不修改论文内容。
/// </summary>
public interface IChatService
{
    /// <summary>
    /// 回答用户的格式问题。
    /// </summary>
    /// <param name="question">用户的问题</param>
    /// <returns>格式相关的回答</returns>
    Task<string> AskAsync(string question, CancellationToken cancellationToken = default);
}
