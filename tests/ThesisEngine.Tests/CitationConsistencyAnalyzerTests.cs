using System.IO.Compression;
using System.Reflection;
using System.Text;
using ThesisEngine.Core;
using ThesisEngine.Core.Citation;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class CitationConsistencyAnalyzerTests
{
    private readonly CitationConsistencyAnalyzer _analyzer = new();

    [Fact]
    public void AcceptsContinuousCitationsWithMatchingJournalReference()
    {
        var result = Analyze(
            "研究结论见 [1]。",
            "参考文献",
            "[1] Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10.");

        Assert.Empty(result.Findings);
        Assert.Equal(1, result.CitationCount);
        Assert.Equal(1, result.ReferenceCount);
    }

    [Fact]
    public void ReportsSkippedFirstUseMissingReferenceAndUncitedEntry()
    {
        var result = Analyze(
            "结论见 [2]。",
            "参考文献",
            "[1] Uncited book. 2022.",
            "[3] Cited book. 2023.");

        Assert.Contains(result.Findings, finding => finding.RuleId == "citation/sequence");
        Assert.Contains(result.Findings, finding => finding.RuleId == "citation/missing-reference");
        Assert.Contains(result.Findings, finding => finding.RuleId == "citation/uncited-reference");
    }

    [Fact]
    public void ReportsCitationPlacedAfterPunctuation()
    {
        var result = Analyze(
            "结论已经说明。[1]",
            "参考文献",
            "[1] Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10.");

        Assert.Contains(result.Findings, finding => finding.RuleId == "citation/punctuation-position");
    }

    [Fact]
    public void ReportsUnnumberedReferenceEntryInsteadOfSilentlyIgnoringIt()
    {
        var result = Analyze(
            "结论见 [1]。",
            "参考文献",
            "Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10.");

        Assert.Contains(result.Findings, finding => finding.RuleId == "citation/reference-number-format");
    }

    [Fact]
    public void ReportsMissingRequiredFieldsWithoutCheckingCitationStyle()
    {
        var result = Analyze(
            "结论见 [1]。",
            "参考文献",
            "[1] Smith. Article title[J]. Journal.");

        Assert.Contains(result.Findings, item => item.RuleId == "citation/reference-fields" && item.Message.Contains("年份"));
        Assert.Contains("卷/期", result.Findings.Single(item => item.RuleId == "citation/reference-fields" && item.Message.Contains("卷/期")).Message);
        Assert.Contains("页码", result.Findings.Single(item => item.RuleId == "citation/reference-fields" && item.Message.Contains("页码")).Message);
    }

    [Theory]
    [InlineData("张三、李四、王五、赵六. 研究题名[J]. 期刊, 2024, 12(2): 3-10.")]
    [InlineData("Smith, A., Jones, B., Brown, C., Green, D. Article title[J]. Journal, 2024, 12(2): 3-10.")]
    public void RequiresEtAlMarkerWhenMoreThanThreeAuthors(string reference)
    {
        var result = Analyze("研究见 [1]。", "参考文献", "[1] " + reference);

        Assert.Contains(result.Findings, finding => finding.RuleId == "citation/reference-fields" &&
            finding.Message.Contains("超过三位作者"));
    }

    [Fact]
    public void ReportsMissingAuthorWhenReferenceStartsWithTitle()
    {
        var result = Analyze(
            "研究见 [1]。",
            "参考文献",
            "[1] 一种研究方法[J]. 期刊, 2024, 12(2): 3-10.");

        Assert.Contains(result.Findings, finding => finding.RuleId == "citation/reference-fields" &&
            finding.Message.Contains("缺少可识别的作者"));
    }

    [Fact]
    public void ReportsSameWorkListedUnderDifferentNumbers()
    {
        var result = Analyze(
            "结论见 [1]。另见 [2]。",
            "参考文献",
            "[1] Smith, A. Same article[J]. Journal, 2024, 12(2): 3-10.",
            "[2] Smith, A. Same article[J]. Journal, 2024, 12(2): 3-10.");

        Assert.Contains(result.Findings, finding => finding.RuleId == "citation/duplicate-work");
    }

    [Theory]
    [InlineData("研究结论见 [1]。", 1, 0)]
    [InlineData("研究结论已经说明。[1]", 1, 1)]
    [InlineData("研究结论见 [1]。", 2, 2)]
    public void SyntheticAcceptanceSamplesMatchExpectedFindingCount(string citationText, int referenceCount, int expectedFindings)
    {
        var paragraphs = new List<CitationParagraph>
        {
            Paragraph(0, citationText),
            Paragraph(1, "参考文献"),
            Paragraph(2, "[1] Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10.")
        };
        if (referenceCount == 2)
            paragraphs.Add(Paragraph(3, "[2] Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10."));

        var result = _analyzer.Analyze(new CitationDocument(paragraphs, 0));

        Assert.Equal(expectedFindings, result.Findings.Count);
    }

    [Fact]
    public void IgnoresOrdinaryArrayIndexButRecognizesSuperscriptCitation()
    {
        var document = new CitationDocument(new[]
        {
            Paragraph(0, "数组 a[1] 的下标表示第一个元素。"),
            Paragraph(1, "研究结果[1]显示", (4, 1, true), (5, 1, true), (6, 1, true)),
            Paragraph(2, "参考文献"),
            Paragraph(3, "[1] Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10.")
        }, 0);

        var result = _analyzer.Analyze(document);

        Assert.Equal(1, result.CitationCount);
        Assert.DoesNotContain(result.Findings, finding => finding.ParagraphIndex == 0);
    }

    [Fact]
    public void RecognizesInlinePlainCitationAndExcludesParameterAndArrayExamples()
    {
        var document = new CitationDocument(new[]
        {
            Paragraph(0, "相关研究[1]支持这一结论。"),
            Paragraph(1, "符号 [1] (参数说明、数组下标) 表示数组元素。"),
            Paragraph(2, "基于索引方法见文献[1]。"),
            Paragraph(3, "参考文献"),
            Paragraph(4, "[1] Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10.")
        }, 0);

        var result = _analyzer.Analyze(document);

        Assert.Equal(2, result.CitationCount);
        Assert.DoesNotContain(result.Findings, finding => finding.ParagraphIndex == 1);
        Assert.DoesNotContain(result.Findings, finding => finding.RuleId == "citation/missing-reference");
    }

    [Fact]
    public void DoesNotTreatUnlistedParameterExampleAsAnUnmatchedCitation()
    {
        var document = new CitationDocument(new[]
        {
            Paragraph(0, "符号 [1] (参数说明、数组下标) 表示数组元素。"),
            Paragraph(1, "研究见 [2]。"),
            Paragraph(2, "参考文献"),
            Paragraph(3, "[2] Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10.")
        }, 0);

        var result = _analyzer.Analyze(document);

        Assert.Equal(1, result.CitationCount);
        Assert.DoesNotContain(result.Findings, finding => finding.RuleId == "citation/missing-reference");
    }

    [Fact]
    public void CliReadsDocxWithoutChangingItAndReportsParagraphLocation()
    {
        var path = CreateDocx(
            "研究结论见 [1]。",
            "参考文献",
            "[1] Smith. Article title[J]. Journal.");
        try
        {
            var before = File.ReadAllBytes(path);
            var output = new StringWriter();
            var error = new StringWriter();

            var command = CreateAnalyzeCommand(output, error);
            var execute = command.GetType().GetMethod(
                "Execute", new[] { typeof(string[]), typeof(TextWriter), typeof(TextWriter) })!;
            var exitCode = (int)execute.Invoke(command,
                new object[] { new[] { "analyze", "--citations", path }, output, error })!;

            Assert.Equal(0, exitCode);
            Assert.Empty(error.ToString());
            Assert.Contains("问题数量：", output.ToString());
            Assert.Contains("word/document.xml#paragraph[2]", output.ToString());
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ReadsSuperscriptCitationFromDocxRunProperties()
    {
        var path = CreateSuperscriptDocx();
        try
        {
            var document = new CitationDocumentReader().Read(path);
            var result = _analyzer.Analyze(document);

            Assert.Contains(document.Paragraphs[0].Runs, run => run.IsSuperscript && run.Length == 3);
            Assert.Equal(1, result.CitationCount);
            Assert.DoesNotContain(result.Findings, finding => finding.RuleId == "citation/missing-reference");
        }
        finally
        {
            File.Delete(path);
        }
    }

    private CitationAnalysisResult Analyze(params string[] paragraphs) =>
        _analyzer.Analyze(new CitationDocument(paragraphs.Select((text, index) => Paragraph(index, text)).ToArray(), 0));

    private static CitationParagraph Paragraph(int index, string text, params (int Start, int Length, bool Superscript)[] runs) =>
        new(index, text, runs.Select(run => new CitationRunSpan(run.Start, run.Length, run.Superscript)).ToArray());

    private static object CreateAnalyzeCommand(TextWriter output, TextWriter error)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln")))
            root = root.Parent;
        Assert.NotNull(root);

        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
        var cliPath = Path.Combine(root!.FullName, "src", "ThesisEngine.Cli", "bin", configuration,
            "net10.0", "ThesisEngine.Cli.dll");
        Assert.True(File.Exists(cliPath), $"CLI was not built: {cliPath}");
        var assembly = Assembly.LoadFrom(cliPath);
        var commandType = assembly.GetType("ThesisEngine.Cli.AnalyzeCommand")!;
        return Activator.CreateInstance(commandType, new UnusedDemoAnalysisService(), output, error)!;
    }

    private sealed class UnusedDemoAnalysisService : IDemoAnalysisService
    {
        public AnalysisResult Analyze(string path, DemoRuleSet profile) =>
            throw new InvalidOperationException("The citation command must bypass demo analysis.");
    }

    private static string CreateDocx(params string[] paragraphs)
    {
        var body = string.Concat(paragraphs.Select(paragraph =>
            "<w:p><w:r><w:t xml:space=\"preserve\">" + System.Security.SecurityElement.Escape(paragraph) + "</w:t></w:r></w:p>"));
        return CreateDocxFromBody(body);
    }

    private static string CreateSuperscriptDocx() => CreateDocxFromBody(
        "<w:p><w:r><w:t>研究结论见 </w:t></w:r><w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:t>[1]</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>参考文献</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>[1] Smith, A. Article title[J]. Journal, 2024, 12(2): 3-10.</w:t></w:r></w:p>");

    private static string CreateDocxFromBody(string body)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        var entry = archive.CreateEntry("word/document.xml");
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>");
        writer.Write(body);
        writer.Write("</w:body></w:document>");
        return path;
    }
}
