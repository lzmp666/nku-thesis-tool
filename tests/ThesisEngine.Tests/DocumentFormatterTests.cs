using System.IO.Compression;
using System.Xml.Linq;
using ThesisEngine.Core.Document;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// #112 第二档：只动 sectPr 的自动修复。
///
/// 这里的重点不是「改对了几个数」，是「没改到不该改的地方」。
/// 每条测试都拿原件和输出逐项比对：包条目顺序、其它部件字节、
/// 去掉 sectPr 之后的正文树、w:t 文本序列、sectPr 里其它元素。
/// </summary>
public sealed class DocumentFormatterTests : IDisposable
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly List<string> _tempFiles = new();

    [Theory]
    [InlineData("document_margin_wrong.docx", 2, "D2")]
    [InlineData("document_paper_wrong.docx", 2, "D1")]
    [InlineData("document_hf_distance_wrong.docx", 2, "D3")]
    [InlineData("document_gutter_wrong.docx", 1, "D4")]
    public void OnlyTheDeviatingAttributesAreWritten(string sample, int expectedFixes, string expectedCode)
    {
        var (input, output, result) = Run(sample);

        Assert.Equal(expectedFixes, result.Fixes.Count);
        Assert.All(result.Fixes, fix => Assert.Equal(expectedCode, fix.Code));
        AssertOnlySectPrChanged(input, output);
    }

    [Fact]
    public void MarginFixUsesClauseDerivedTwips()
    {
        var (_, _, result) = Run("document_margin_wrong.docx");

        Assert.Contains(result.Fixes, fix => fix.Description.Contains("上页边距 1417 → 2154 twips"));
        Assert.Contains(result.Fixes, fix => fix.Description.Contains("下页边距 1417 → 2154 twips"));
    }

    [Fact]
    public void CompliantSampleNeedsNoChanges()
    {
        var (input, output, result) = Run("document_baseline_correct.docx");

        Assert.Empty(result.Fixes);
        Assert.Empty(result.Skips);
        AssertOnlySectPrChanged(input, output);
    }

    /// <summary>复诊：改完 D1–D4 必须清零，不能出现「查出 5 条只修 3 条」。</summary>
    [Theory]
    [InlineData("document_margin_wrong.docx")]
    [InlineData("document_paper_wrong.docx")]
    [InlineData("document_hf_distance_wrong.docx")]
    [InlineData("document_gutter_wrong.docx")]
    [InlineData("document_sections_same_numbering.docx")]
    public void AutoFixableIssuesAreClearedAfterWrite(string sample)
    {
        var (_, output, _) = Run(sample);
        var after = new DocumentAnalyzer().Analyze(output);

        Assert.DoesNotContain(after.Issues,
            issue => issue.Code is "D1" or "D2" or "D3" or "D4");
    }

    /// <summary>issue 明确要求只报告不改的三类，修复后问题必须还在。</summary>
    [Theory]
    [InlineData("document_toc_manual.docx", "D9")]
    [InlineData("document_blank_page.docx", "D10")]
    [InlineData("document_footer_text.docx", "D7")]
    [InlineData("document_header_font_wrong.docx", "D11")]
    [InlineData("document_footer_size_wrong.docx", "D12")]
    public void ReportOnlyIssuesSurviveTheWrite(string sample, string code)
    {
        var (_, output, _) = Run(sample);
        var after = new DocumentAnalyzer().Analyze(output);

        Assert.Contains(after.Issues, issue => issue.Code == code);
    }

    [Fact]
    public void AllSectionsAreFixed()
    {
        var (input, output, result) = Run("document_sections_same_numbering.docx");

        Assert.Equal(2, result.SectionCount);
        AssertOnlySectPrChanged(input, output);
    }

    [Fact]
    public void SecondRunIsByteForByteIdentical()
    {
        var (_, first, _) = Run("document_margin_wrong.docx");
        var secondPath = Temp();
        var second = new DocumentFormatter().Apply(first, secondPath);

        Assert.Empty(second.Fixes);
        Assert.Equal(File.ReadAllBytes(first), File.ReadAllBytes(secondPath));
    }

    [Fact]
    public void MissingPageSizeElementIsSkippedRatherThanInserted()
    {
        var input = CreateMinimalDocx("<w:sectPr/>");

        var result = new DocumentFormatter().Apply(input, Temp());

        Assert.Empty(result.Fixes);
        Assert.Contains(result.Skips, skip => skip.Contains("没有 pgSz"));
        Assert.Contains(result.Skips, skip => skip.Contains("没有 pgMar"));
    }

    [Fact]
    public void MissingAttributesAreFilledWithoutAddingElements()
    {
        var input = CreateMinimalDocx(
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1417\"/><w:titlePg/></w:sectPr>");
        var outputPath = Temp();

        var result = new DocumentFormatter().Apply(input, outputPath);

        Assert.Equal(9, result.Fixes.Count); // pgSz 2 项 + pgMar 7 项
        var sectPr = ReadDocument(outputPath).Descendants(W + "sectPr").Single();
        Assert.Single(sectPr.Elements(W + "titlePg"));
        Assert.Equal("11906", sectPr.Element(W + "pgSz")!.Attribute(W + "w")!.Value);
        Assert.Equal("2154", sectPr.Element(W + "pgMar")!.Attribute(W + "top")!.Value);
        Assert.Equal("1701", sectPr.Element(W + "pgMar")!.Attribute(W + "header")!.Value);
    }

    [Fact]
    public void SameInputAndOutputPathIsRejected()
    {
        var input = SamplePath("document_baseline_correct.docx");

        Assert.Throws<ArgumentException>(() => new DocumentFormatter().Apply(input, input));
    }

    /// <summary>硬边界：除 sectPr 里的 pgSz / pgMar 属性外，一个字节都不许变。</summary>
    private static void AssertOnlySectPrChanged(string inputPath, string outputPath)
    {
        var before = ReadParts(inputPath);
        var after = ReadParts(outputPath);

        Assert.Equal(before.Keys.ToList(), after.Keys.ToList());

        foreach (var name in before.Keys.Where(name => name != "word/document.xml"))
        {
            Assert.Equal(before[name], after[name]);
        }

        var beforeXml = before["word/document.xml"];
        var afterXml = after["word/document.xml"];

        Assert.True(XNode.DeepEquals(WithoutSectPr(beforeXml), WithoutSectPr(afterXml)),
            "去掉 sectPr 之后，正文 XML 树必须完全一致");
        Assert.Equal(Texts(beforeXml), Texts(afterXml));
        Assert.Equal(OtherSectElements(beforeXml), OtherSectElements(afterXml));
    }

    private static Dictionary<string, byte[]> ReadParts(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        });
    }

    private static XDocument WithoutSectPr(byte[] xml)
    {
        var document = XDocument.Load(new MemoryStream(xml));
        document.Descendants(W + "sectPr").Remove();
        return document;
    }

    private static List<string> Texts(byte[] xml) =>
        XDocument.Load(new MemoryStream(xml))
            .Descendants(W + "t")
            .Select(node => node.Value)
            .ToList();

    private static List<string> OtherSectElements(byte[] xml) =>
        XDocument.Load(new MemoryStream(xml))
            .Descendants(W + "sectPr")
            .SelectMany(section => section.Elements()
                .Where(element => element.Name != W + "pgSz" && element.Name != W + "pgMar")
                .Select(element => element.ToString(SaveOptions.DisableFormatting)))
            .ToList();

    private static XDocument ReadDocument(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
        return XDocument.Load(reader);
    }

    private (string Input, string Output, DocumentFormatResult Result) Run(string sample)
    {
        var input = SamplePath(sample);
        Assert.True(File.Exists(input), $"样本不存在：{input}");
        var output = Temp();
        var result = new DocumentFormatter().Apply(input, output);
        return (input, output, result);
    }

    private string CreateMinimalDocx(string sectPrXml)
    {
        var path = Temp();
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        using var writer = new StreamWriter(archive.CreateEntry("word/document.xml").Open());
        writer.Write(
            $"<w:document xmlns:w=\"{W}\"><w:body>" +
            "<w:p><w:r><w:t>合成样本正文</w:t></w:r></w:p>" +
            sectPrXml + "</w:body></w:document>");
        return path;
    }

    private string Temp()
    {
        var path = Path.Combine(Path.GetTempPath(), "nku-document-" + Guid.NewGuid().ToString("N") + ".docx");
        _tempFiles.Add(path);
        return path;
    }

    private static string SamplePath(string sampleName)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return Path.Combine(directory!.FullName, "tests", "fixtures", "synthetic", "document", sampleName);
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            if (File.Exists(file)) File.Delete(file);
        }
    }
}
