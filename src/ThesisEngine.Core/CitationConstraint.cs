namespace ThesisEngine.Core;

/// <summary>
/// 引用 / 标注硬约束：参考文献必须符合 GB/T 7714。做不到（格式不合规）就明确报错，不允许含糊带过。
/// </summary>
public static class CitationConstraint
{
    /// <summary>判断问题是否属于参考文献 / 引用 / 标注相关。</summary>
    public static bool IsCitationQuestion(string question)
    {
        var q = (question ?? string.Empty).Trim();
        return q.Contains("参考文献") || q.Contains("引用") || q.Contains("引文") ||
               q.Contains("脚注") || q.Contains("标注") || q.Contains("GB/T 7714") ||
               q.Contains("著录") || q.Contains("文献");
    }

    /// <summary>
    /// 校验一条参考文献著录是否符合 GB/T 7714 的基本结构。
    /// 合规返回 null；否则返回具体、不含糊的不合规原因。
    /// </summary>
    public static string? ValidateCitation(string citation)
    {
        var c = (citation ?? string.Empty).Trim();
        if (c.Length == 0)
            return "引用内容为空，无法校验。";
        if (!c.Contains('.'))
            return "缺少必要的标点“.”：GB/T 7714 要求作者、题名、出版信息之间用“.”分隔。";
        if (!c.Contains('['))
            return "缺少文献类型标识：需用 [J]（期刊）、[M]（专著）、[D]（学位论文）等标注。";
        if (!HasFourDigitYear(c))
            return "缺少出版年：GB/T 7714 要求标注四位数字的出版年。";
        return null;
    }

    private static bool HasFourDigitYear(string text)
    {
        for (var i = 0; i + 4 <= text.Length; i++)
        {
            if (char.IsDigit(text[i]) && char.IsDigit(text[i + 1]) &&
                char.IsDigit(text[i + 2]) && char.IsDigit(text[i + 3]))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>引用问题的权威回答（硬约束说明）。</summary>
    public static string CanonicalAnswer =>
        "参考文献必须符合 GB/T 7714 著录格式，例如：作者. 题名[文献类型标识]. 出版地: 出版者, 出版年. " +
        "正文中用上标[1]标注序号，并与文末列表一一对应；不得缺失作者、题名、文献类型、出版年和出版者。";
}
