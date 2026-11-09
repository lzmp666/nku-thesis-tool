using System.IO.Compression;
using ThesisEngine.Core.Formula;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// #114 第一档：公式排版检查。
///
/// 样本在 tests/fixtures/synthetic/formula/，每个文件只偏离一个维度，
/// 由 generate.py 生成、manifest.json 登记预期。除样本外还手工造两个反例，
/// 专门守住「纯 OMML 段没有 w:r 也不能漏检」和「行内公式不能当成独立公式」。
/// </summary>
public sealed class FormulaAnalyzerTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Fact]
    public void BaselineSampleHasNoIssuesAndSplitsInlineFromStandalone()
    {
        var result = Analyze("formula_baseline_correct.docx");

        Assert.Equal(2, result.Blocks.Count);
        Assert.Equal(2, result.InlineCount);
        Assert.Empty(result.Issues);
        // 两种居中写法都要认：一个靠 pPr/jc，一个靠 m:oMathParaPr/jc。
        Assert.All(result.Blocks, block => Assert.Equal(FormulaLayout.Standalone, block.Layout));
        Assert.Equal("(1)", result.Blocks[0].EquationNumber);
        Assert.Equal("(2)", result.Blocks[1].EquationNumber);
    }

    [Fact]
    public void NotCenteredSampleIsReportedAsConventionNotAsClause()
    {
        var result = Analyze("formula_not_centered.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("F1", issue.Code);
        // 规范里没有「公式居中」这一条，就不能算成不符合规范。
        Assert.False(issue.HasClause);
        Assert.Empty(result.ClauseIssues);
    }

    [Fact]
    public void NumberNotRightAlignedSampleIsReportedWithClauseBasis()
    {
        var result = Analyze("formula_number_not_right_aligned.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("F2", issue.Code);
        Assert.True(issue.HasClause);
        Assert.Contains("R13", issue.Basis);
    }

    [Fact]
    public void NumberFormatWrongSampleReportsBracketFontAndSize()
    {
        var result = Analyze("formula_number_format_wrong.docx");

        var codes = result.Issues.Select(i => i.Title).ToList();
        Assert.Equal(3, codes.Count);
        Assert.Contains("式号未加圆括号", codes);
        Assert.Contains("式号西文字体不对", codes);
        Assert.Contains("式号字号不对", codes);
        Assert.All(result.Issues, issue => Assert.True(issue.HasClause));
    }

    [Fact]
    public void SpacingWrongSampleIsReported()
    {
        var result = Analyze("formula_spacing_wrong.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("F5", issue.Code);
        Assert.Contains("R11", issue.Basis);
    }

    [Fact]
    public void TableRowLayoutIsAccepted()
    {
        var result = Analyze("formula_table_row_correct.docx");

        var block = Assert.Single(result.Blocks);
        Assert.Equal(FormulaLayout.TableRow, block.Layout);
        Assert.Equal("(3)", block.EquationNumber);
        // 表格排版下位置由列宽决定，不能报「未居中」。
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void InlineOnlySampleHasNoStandaloneFormula()
    {
        var result = Analyze("formula_inline_only.docx");

        Assert.Empty(result.Blocks);
        Assert.Equal(2, result.InlineCount);
        Assert.Empty(result.Issues);
    }

    /// <summary>
    /// 关键回归：纯 OMML 的公式段没有 w:r、没有 w:t，任何「先取 run 或段落文本再判断」
    /// 的写法都会跳过它，结果是公式全不合规、报告却是 0 条。这里造一个只有 oMathPara 的段落，
    /// 断言它必须被识别出来并报出未居中。
    /// </summary>
    [Fact]
    public void StandaloneFormulaWithoutWordRunsIsStillDetected()
    {
        const string body = """
            <w:p><w:pPr><w:jc w:val="left"/></w:pPr>
              <m:oMathPara><m:oMath>
                <m:r><m:t>E</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>m</m:t></m:r>
              </m:oMath></m:oMathPara>
            </w:p>
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new FormulaAnalyzer().Analyze(path);

            var block = Assert.Single(result.Blocks);
            Assert.Equal(FormulaLayout.Standalone, block.Layout);
            Assert.False(block.HasEquationNumber);
            var issue = Assert.Single(result.Issues);
            Assert.Equal("F1", issue.Code);
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 反向用例：句子中间的公式不是独立公式段，不能被套上「未居中 / 未右对齐」。
    /// </summary>
    [Fact]
    public void InlineFormulaInsideSentenceIsNotTreatedAsStandalone()
    {
        const string body = """
            <w:p><w:pPr><w:jc w:val="both"/></w:pPr>
              <w:r><w:t>当</w:t></w:r>
              <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
              <w:r><w:t>取最大值时结论成立。</w:t></w:r>
            </w:p>
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new FormulaAnalyzer().Analyze(path);

            Assert.Empty(result.Blocks);
            Assert.Equal(1, result.InlineCount);
            Assert.Empty(result.Issues);
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 裸数字式号得能认出来，否则「未加圆括号」这条永远报不出来——
    /// 认不到式号就当成「无式号」静静放过，是最难发现的那类假阴性。
    /// </summary>
    [Fact]
    public void BareNumberIsRecognizedAsEquationNumber()
    {
        const string body = """
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
              <m:oMathPara><m:oMath><m:r><m:t>E</m:t></m:r></m:oMath></m:oMathPara>
              <w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="21"/></w:rPr><w:t>3</w:t></w:r>
            </w:p>
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new FormulaAnalyzer().Analyze(path);

            var block = Assert.Single(result.Blocks);
            Assert.True(block.HasEquationNumber);
            Assert.Equal("3", block.EquationNumber);
            Assert.Contains(result.Issues, issue => issue.Title == "式号未加圆括号");
            // 字体字号都对，不该顺手多报两条。
            Assert.DoesNotContain(result.Issues, issue => issue.Title == "式号西文字体不对");
            Assert.DoesNotContain(result.Issues, issue => issue.Title == "式号字号不对");
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// Word 16 把式号抄在公式对象里（<c>&lt;m:r&gt;&lt;w:tab/&gt;&lt;m:t&gt;(1)&lt;/m:t&gt;&lt;/m:r&gt;</c>），
    /// 是 m:t 不是 w:t。只读 w:t 的实现会认不出式号，F2/F3 从此永不检查。
    /// 这个结构是在真机 Word 16.0 上实测出来的，不是推测。
    /// </summary>
    [Fact]
    public void WordNativeSampleRecognizesNumberInsideMathObject()
    {
        var result = Analyze("formula_word_native_correct.docx");

        var block = Assert.Single(result.Blocks);
        Assert.True(block.HasEquationNumber);
        Assert.Equal("(1)", block.EquationNumber);
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void WordNativeSampleWithoutTabStillReportsAlignment()
    {
        var result = Analyze("formula_word_native_number_not_aligned.docx");

        var block = Assert.Single(result.Blocks);
        Assert.Equal("(1)", block.EquationNumber);
        Assert.Contains(result.Issues, issue => issue.Code == "F2" && issue.HasClause);
        Assert.Contains(result.Issues, issue => issue.Code == "F1");
    }

    /// <summary>
    /// 反向用例：公式末尾是普通数字（E=mc² 的 2）时，不能把它当式号，
    /// 否则会凭空报一条「式号未加圆括号」。
    /// </summary>
    [Fact]
    public void TrailingMathDigitIsNotMistakenForEquationNumber()
    {
        const string body = """
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
              <m:oMathPara><m:oMath>
                <m:r><m:t>E=mc</m:t></m:r>
                <m:r><m:t>2</m:t></m:r>
              </m:oMath></m:oMathPara>
            </w:p>
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new FormulaAnalyzer().Analyze(path);

            var block = Assert.Single(result.Blocks);
            Assert.False(block.HasEquationNumber);
            Assert.Empty(result.Issues);
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 式号在公式对象里时，字体是公式引擎给的（实测是 Cambria Math），不是用户选的，
    /// 所以不报字体；但字号仍然是用户能改、规范也明确要求 10.5 磅的，必须报。
    /// </summary>
    [Fact]
    public void NumberInsideMathSkipsFontButStillChecksSize()
    {
        const string body = """
            <w:p><w:pPr><w:jc w:val="center"/></w:pPr>
              <m:oMathPara><m:oMath>
                <m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr><m:t>E</m:t></m:r>
                <m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/><w:sz w:val="24"/></w:rPr><m:t>(1)</m:t></m:r>
              </m:oMath></m:oMathPara>
            </w:p>
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new FormulaAnalyzer().Analyze(path);

            Assert.Contains(result.Issues, issue => issue.Title == "式号字号不对");
            Assert.DoesNotContain(result.Issues, issue => issue.Title == "式号西文字体不对");
        }
        finally { File.Delete(path); }
    }

    private static FormulaAnalysisResult Analyze(string sampleName)
    {
        var root = FindRepositoryRoot();
        var path = Path.Combine(root, "tests", "fixtures", "synthetic", "formula", sampleName);
        Assert.True(File.Exists(path), $"样本不存在：{path}");
        return new FormulaAnalyzer().Analyze(path);
    }

    private static string CreateDocx(string bodyXml)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml",
            """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>""");
        Add(archive, "_rels/.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""");
        Add(archive, "word/document.xml",
            $"""<w:document xmlns:w="{W}" xmlns:m="{M}"><w:body>{bodyXml}</w:body></w:document>""");
        return path;
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(content);
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }
}
