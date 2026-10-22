using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class NkuGraduate2026BodyAnalyzerTests
{
    private static readonly NkuGraduate2026BodyProfile Profile = NkuGraduate2026BodyProfile.Draft;

    [Fact]
    public void CorrectMixedScriptBodyTextProducesNoDiagnostics()
    {
        var result = Analyze(Run("中文 Latin", "宋体", "Times New Roman", "Times New Roman", 12));

        Assert.Empty(result.Diagnostics);
        Assert.Equal("nku-graduate-2026-body-v1", result.RuleVersion);
        Assert.Contains("draft", result.Coverage);
        Assert.Contains("1 个已确认正文文本片段", result.Coverage);
    }

    [Fact]
    public void IndependentlyChecksHanLatinSlotsAndSize()
    {
        var result = Analyze(Run("中文 Latin", "黑体", "Arial", "Calibri", 10.5));

        Assert.Collection(result.Diagnostics,
            d => Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdFontSize, d.RuleId),
            d => Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdChineseFont, d.RuleId),
            d => Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdLatinFont, d.RuleId));
        Assert.All(result.Diagnostics, d =>
        {
            Assert.Equal("manual", d.FixLevel);
            Assert.Contains("M26-036", d.Message);
            Assert.Contains("规则仍为草案", d.Message);
        });
    }

    [Fact]
    public void OnlyChecksFontSlotsForScriptsPresentInTheText()
    {
        var han = Analyze(Run("只有中文", "黑体", null, null, 12));
        var latin = Analyze(Run("Latin only", null, "Arial", "Arial", 12));
        var digits = Analyze(Run("2026，. + √", null, "Arial", "Arial", 12));

        Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdChineseFont, Assert.Single(han.Diagnostics).RuleId);
        Assert.Empty(han.UnsupportedScopes);
        Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdLatinFont, Assert.Single(latin.Diagnostics).RuleId);
        Assert.Empty(latin.UnsupportedScopes);
        Assert.Empty(digits.Diagnostics);
        Assert.Empty(digits.UnsupportedScopes);
    }

    [Fact]
    public void UnknownScopeAndEffectiveValuesAreUnsupportedNotViolations()
    {
        var paragraph = new ParagraphSnapshot(0,
            new[] { new TextRunSnapshot(0, "黑体", "Arial", 10, "正文", "Arial", "Arial") },
            ParagraphScopeClassification.Unknown, "边界不明");
        var result = new NkuGraduate2026BodyAnalyzer().Analyze(
            new DocumentSnapshot("word/document.xml", new[] { paragraph }, Array.Empty<string>()), Profile);

        Assert.Empty(result.Diagnostics);
        Assert.Contains(result.UnsupportedScopes, x => x.Contains("paragraph[0]") && x.Contains("边界不明"));
    }

    [Fact]
    public void ExcludedParagraphsDoNotAffectCoverageOrDiagnostics()
    {
        var paragraph = new ParagraphSnapshot(0,
            new[] { new TextRunSnapshot(0, "黑体", "Arial", 10, "参考文献", "Arial", "Arial") },
            ParagraphScopeClassification.Excluded, "标题或非正文部件");
        var result = new NkuGraduate2026BodyAnalyzer().Analyze(
            new DocumentSnapshot("word/document.xml", new[] { paragraph }, Array.Empty<string>()), Profile);

        Assert.Empty(result.Diagnostics);
        Assert.Empty(result.UnsupportedScopes);
        Assert.Contains("已检查 0 个", result.Coverage);
    }

    [Fact]
    public void UnresolvedStyleChainIsUnsupportedWithoutInspectingRunValues()
    {
        var result = Analyze(Run("中文 Latin", "黑体", "Arial", "Arial", 10, unresolvedStyle: true));

        Assert.Empty(result.Diagnostics);
        Assert.Contains(result.UnsupportedScopes, x => x.Contains("有效样式链无法解析"));
    }

    [Fact]
    public void MissingEffectiveValuesRemainUnsupportedRatherThanViolations()
    {
        var result = Analyze(Run("中文 Latin", null, null, null, null));

        Assert.Empty(result.Diagnostics);
        Assert.Equal(3, result.UnsupportedScopes.Count);
        Assert.Contains(result.UnsupportedScopes, x => x.Contains("字号有效值未知"));
        Assert.Contains(result.UnsupportedScopes, x => x.Contains("eastAsia 字体槽有效值未知"));
        Assert.Contains(result.UnsupportedScopes, x => x.Contains("ascii 与 hAnsi"));
    }

    [Fact]
    public void ServiceUsesInheritedBodyFormattingAndDoesNotApplyRulesToOtherParts()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
              <w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
              <w:style w:type="paragraph" w:styleId="BodyParent"><w:rPr><w:rFonts w:eastAsia="黑体" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="21"/></w:rPr></w:style>
              <w:style w:type="paragraph" w:styleId="BodyChild"><w:basedOn w:val="BodyParent"/></w:style>
              <w:style w:type="character" w:styleId="GoodAll"><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:style>
            </w:styles>
            """;
        var body = """
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>摘要</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="BodyChild"/></w:pPr><w:r><w:t>摘要 Latin</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>目录</w:t></w:r></w:p>
            <w:p><w:fldSimple w:instr="TOC"><w:r><w:t>目录内容</w:t></w:r></w:fldSimple></w:p>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章绪论</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="BodyChild"/></w:pPr><w:r><w:t>正文 Latin</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="BodyChild"/></w:pPr><w:r><w:rPr><w:rStyle w:val="GoodAll"/></w:rPr><w:t>正确 Latin</w:t></w:r></w:p>
            <w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="BodyChild"/></w:pPr><w:r><w:t>表格 Latin</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="BodyChild"/></w:pPr><w:r><w:t>文献 Latin</w:t></w:r></w:p>
            """;
        var path = CreateDocx(body, styles, ("word/footnotes.xml", "<footnotes/>"), ("word/header1.xml", "<header/>"));
        try
        {
            var before = SHA256.HashData(File.ReadAllBytes(path));
            var result = new NkuGraduate2026BodyAnalysisService().Analyze(path, Profile);

            Assert.Equal(before, SHA256.HashData(File.ReadAllBytes(path)));
            Assert.Equal(Profile.RuleVersion, result.RuleVersion);
            Assert.Collection(result.Diagnostics,
                d => { Assert.Equal("NKU-REQ-044", d.RuleId); Assert.Equal("12pt", d.Expected); Assert.Equal("10.5pt", d.Actual); },
                d => { Assert.Equal("NKU-REQ-042", d.RuleId); Assert.Equal("宋体", d.Expected); Assert.Equal("黑体", d.Actual); },
                d => { Assert.Equal("NKU-REQ-043", d.RuleId); Assert.Equal("ascii=Times New Roman; hAnsi=Times New Roman", d.Expected); Assert.Equal("ascii=Arial; hAnsi=Arial", d.Actual); });
            Assert.All(result.Diagnostics, d =>
            {
                Assert.Equal("word/document.xml#paragraph[5]/run[0]", d.Location);
                Assert.Equal("manual", d.FixLevel);
                Assert.Equal("warning", d.Severity);
                Assert.Contains("2026版", d.Message);
                Assert.Contains("第4.7节表4.6（M26-036）", d.Message);
            });
            Assert.Contains(result.UnsupportedScopes, scope => scope.StartsWith("tables", StringComparison.Ordinal));
            Assert.Contains("footnotes", result.UnsupportedScopes);
            Assert.Contains("headers/footers", result.UnsupportedScopes);
            Assert.Contains("simple fields", result.UnsupportedScopes);
            Assert.Contains("已检查 2 个", result.Coverage);
            AssertGolden(result);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void OrdinaryTextInMathParagraphIsCheckedWhileEquationObjectsRemainUnsupported()
    {
        var body = """
            <w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>第一章 绪论</w:t></w:r></w:p>
            <w:p><w:r><w:rPr><w:rFonts w:eastAsia="黑体" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="21"/></w:rPr><w:t>正文 Latin</w:t></w:r><m:oMath><m:r><m:t>x+1</m:t></m:r></m:oMath></w:p>
            <w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p>
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new NkuGraduate2026BodyAnalysisService().Analyze(path, Profile);

            Assert.Equal(3, result.Diagnostics.Count);
            Assert.All(result.Diagnostics, d => Assert.Equal("word/document.xml#paragraph[1]/run[0]", d.Location));
            Assert.Contains(result.UnsupportedScopes, x => x.Contains("equation objects"));
            Assert.Contains("已检查 1 个", result.Coverage);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void RejectsApprovalUntilProfileAcceptanceTestsAreComplete()
    {
        var profile = Profile with { Status = "approved" };
        var snapshot = new DocumentSnapshot("word/document.xml", Array.Empty<ParagraphSnapshot>(), Array.Empty<string>());

        Assert.Throws<ArgumentException>(() => new NkuGraduate2026BodyAnalyzer().Analyze(snapshot, profile));
    }

    [Fact]
    public void RejectsProfileIdentityOrValuesThatDoNotMatchTheImplementedDraft()
    {
        var snapshot = new DocumentSnapshot("word/document.xml", Array.Empty<ParagraphSnapshot>(), Array.Empty<string>());
        var analyzer = new NkuGraduate2026BodyAnalyzer();

        Assert.Throws<ArgumentException>(() => analyzer.Analyze(snapshot, Profile with { RuleVersion = "nku-graduate-2026-body-v2" }));
        Assert.Throws<ArgumentException>(() => analyzer.Analyze(snapshot, Profile with { LatinFontExpected = "Arial" }));
        Assert.Throws<ArgumentException>(() => analyzer.Analyze(snapshot, Profile with { FontSizeExpectedPt = double.NaN }));
        Assert.Throws<ArgumentException>(() => analyzer.Analyze(snapshot, Profile with { FontSizeExpectedPt = 12.005 }));
    }

    private static AnalysisResult Analyze(TextRunSnapshot run)
    {
        var paragraph = new ParagraphSnapshot(0, new[] { run }, ParagraphScopeClassification.ConfirmedBodyProse);
        var snapshot = new DocumentSnapshot("word/document.xml", new[] { paragraph }, Array.Empty<string>());
        return new NkuGraduate2026BodyAnalyzer().Analyze(snapshot, Profile);
    }

    private static TextRunSnapshot Run(string text, string? eastAsia, string? ascii, string? hAnsi, double? size,
        bool unresolvedStyle = false) =>
        new(0, eastAsia, ascii, size, text, ascii, hAnsi, unresolvedStyle);

    private static string CreateDocx(string body, string? styles = null, params (string Name, string Xml)[] otherParts)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add("word/document.xml", $"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body>{body}<w:sectPr/></w:body></w:document>");
        if (styles is not null) Add("word/styles.xml", styles);
        foreach (var part in otherParts) Add(part.Name, part.Xml);
        return path;

        void Add(string name, string xml)
        {
            using var writer = new StreamWriter(archive.CreateEntry(name).Open());
            writer.Write(xml);
        }
    }

    private static void AssertGolden(AnalysisResult result)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln"))) root = root.Parent;
        Assert.NotNull(root);
        using var expected = JsonDocument.Parse(File.ReadAllText(Path.Combine(root.FullName, "tests", "goldens", "nku-2026-body-v1.json")));
        var actual = JsonSerializer.Serialize(new { result.RuleVersion, result.Diagnostics },
            new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        using var actualJson = JsonDocument.Parse(actual);
        Assert.Equal(JsonSerializer.Serialize(expected.RootElement), JsonSerializer.Serialize(actualJson.RootElement));
    }
}
