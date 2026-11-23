using System.IO.Compression;
using System.Security.Cryptography;
using ThesisEngine.Core;
using Xunit;

/// <summary>
/// draft 规则保守修正器（apply-draft-format）回归测试。
/// 核心保证：① 修后对应分析器问题归零；② 未判定维度不被误改；③ 幂等。
/// </summary>
public sealed class DraftFormatRewriterTests
{
    private readonly DraftFormatRewriter _rewriter = new();

    private static string RepoRoot()
    {
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d is not null && !File.Exists(Path.Combine(d.FullName, "NkuThesisTool.sln")))
            d = d.Parent;
        Assert.NotNull(d);
        return d!.FullName;
    }

    private static string Fixture(string folder, string name) =>
        Path.Combine(RepoRoot(), "tests", "fixtures", "synthetic", folder, name);

    private static string ShaOfDocumentXml(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        using var stream = zip.GetEntry("word/document.xml")!.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return Convert.ToHexString(SHA256.HashData(ms.ToArray()));
    }

    public static TheoryData<string, string, string> WrongSamples
    {
        get
        {
            var data = new TheoryData<string, string, string>();
            data.Add("abstract", "abstract", "abstract_cn_body_font_wrong.docx");
            data.Add("abstract", "abstract", "abstract_en_title_font_wrong.docx");
            data.Add("toc", "toc", "toc_entry_indent_wrong.docx");
            data.Add("back-matter", "back-matter", "back_matter_ack_cn_font_wrong.docx");
            data.Add("body-text", "body-text", "body_line_spacing_wrong.docx");
            data.Add("captions", "captions", "captions_figure_wrong.docx");
            data.Add("headings", "headings", "headings_all_wrong.docx");
            return data;
        }
    }

    private static int IssueCount(string analyzerName, string path) => analyzerName switch
    {
        "abstract" => new AbstractFormatAnalyzer().Analyze(path).Issues.Count,
        "toc" => new TocFormatAnalyzer().Analyze(path).Issues.Count,
        "back-matter" => new BackMatterFormatAnalyzer().Analyze(path).Issues.Count,
        "body-text" => new BodyTextFormatAnalyzer().Analyze(path).Issues.Count,
        "captions" => new CaptionFormatAnalyzer().Analyze(path).Issues.Count,
        "headings" => new HeadingFormatAnalyzer().Analyze(path).Issues.Count,
        _ => throw new ArgumentOutOfRangeException(nameof(analyzerName), analyzerName, null)
    };

    [Theory]
    [MemberData(nameof(WrongSamples))]
    public void FixMakesExplicitIssuesDisappear(string folder, string analyzerCommand, string name)
    {
        var input = Fixture(folder, name);
        var before = IssueCount(analyzerCommand, input);
        Assert.True(before > 0, $"样本应有问题：{name}");

        var output = Path.Combine(Path.GetTempPath(), "draft-fix-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            var result = _rewriter.Apply(input, output);
            Assert.True(result.HasChanges, $"修正器应产生改动：{name}");
            Assert.Equal(0, IssueCount(analyzerCommand, output));
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void FixIsIdempotent()
    {
        var input = Fixture("captions", "captions_figure_wrong.docx");
        var first = Path.Combine(Path.GetTempPath(), "draft-idem1-" + Guid.NewGuid().ToString("N") + ".docx");
        var second = Path.Combine(Path.GetTempPath(), "draft-idem2-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            var r1 = _rewriter.Apply(input, first);
            var r2 = _rewriter.Apply(first, second);
            Assert.True(r1.HasChanges);
            Assert.Equal(0, r2.Applied);
            Assert.Equal(ShaOfDocumentXml(first), ShaOfDocumentXml(second));
        }
        finally
        {
            if (File.Exists(first)) File.Delete(first);
            if (File.Exists(second)) File.Delete(second);
        }
    }

    [Fact]
    public void CrossAnalyzerConflictIsSkippedNotGuessed()
    {
        // 标题样本中正文分析器会把编号标题误纳入正文，两个口径冲突时不得硬改。
        var input = Fixture("headings", "headings_size_wrong.docx");
        var output = Path.Combine(Path.GetTempPath(), "draft-conflict-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            var result = _rewriter.Apply(input, output);
            // 只改了标题本身那一处字号；标题的行距/段距保持原值（未被正文口径覆盖）。
            Assert.Equal(1, result.Applied);
            Assert.Equal(0, IssueCount("headings", output));
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void CompliantSampleReportsNoIssuesForItsOwnAnalyzer()
    {
        Assert.Equal(0, IssueCount("abstract", Fixture("abstract", "abstract_compliant.docx")));
        Assert.Equal(0, IssueCount("toc", Fixture("toc", "toc_compliant.docx")));
        Assert.Equal(0, IssueCount("headings", Fixture("headings", "headings_compliant.docx")));
    }
}
