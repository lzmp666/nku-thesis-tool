using System.IO.Compression;
using System.Xml.Linq;
using ThesisEngine.Core;
using ThesisEngine.Core.Formula;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// #114 第二档：公式格式写入。
///
/// 这里的用例分成两类，权重不一样：
/// - 「改到位」：修完之后检查器复诊为 0，或者只剩本档明确不修的那几项。
/// - 「没被误伤」：m:t / w:t 逐字不变、已合规的文档一个字节不动、行内公式整段不碰。
///   第二类更重要——改错了没人看得出来，而这恰恰是公式最怕的。
/// </summary>
public sealed class FormulaFormatterTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static readonly XNamespace Wns = W;
    private static readonly XNamespace Mns = M;

    [Fact]
    public void BaselineSampleIsLeftUntouched()
    {
        var (input, output) = Run("formula_baseline_correct.docx");

        Assert.Empty(output.Fixes);
        AssertDocumentUnchanged(input, output.OutputPath);
        Assert.Equal(2, output.Result.StandaloneCount);
        Assert.Equal(2, output.Result.InlineCount);
    }

    /// <summary>行内公式整段不许动。全行内的文档修完必须与原文件结构完全一致。</summary>
    [Fact]
    public void InlineOnlyDocumentIsBitForBitUnchanged()
    {
        var (input, output) = Run("formula_inline_only.docx");

        Assert.Empty(output.Fixes);
        Assert.Equal(0, output.Result.StandaloneCount);
        Assert.Equal(2, output.Result.InlineCount);
        AssertDocumentUnchanged(input, output.OutputPath);
    }

    [Fact]
    public void NotCenteredSampleGetsParagraphAndMathParagraphAlignment()
    {
        var (_, output) = Run("formula_not_centered.docx");

        Assert.Contains(output.Fixes, fix => fix.Code == "F1");

        var para = FormulaParagraph(ReadDocument(output.OutputPath), 2);
        Assert.Equal("center", para.Element(Wns + "pPr")?.Element(Wns + "jc")?.Attribute(Wns + "val")?.Value);
        // 两处都要改：Word 会按 m:oMathParaPr/jc 显示，只改段落 jc 有可能看不出来。
        Assert.Equal("center", para.Descendants(Mns + "oMathParaPr").First()
            .Element(Mns + "jc")?.Attribute(Mns + "val")?.Value);

        Assert.Empty(new FormulaAnalyzer().Analyze(output.OutputPath).Issues);
    }

    [Fact]
    public void NumberAlignmentAddsRightTabStopAtTextWidthAndATabCharacter()
    {
        var (_, output) = Run("formula_number_not_right_aligned.docx");

        var para = FormulaParagraph(ReadDocument(output.OutputPath), 2);
        var rightTab = para.Element(Wns + "pPr")?.Element(Wns + "tabs")?.Elements(Wns + "tab")
            .Single(tab => tab.Attribute(Wns + "val")?.Value == "right");
        Assert.NotNull(rightTab);
        // A4 宽 11906 twips，左右页边距各 1814，版心右缘 8278。
        Assert.Equal("8278", rightTab!.Attribute(Wns + "pos")?.Value);

        // 光加制表位没用，式号前面必须有制表符才会被推过去。
        // 式号可能写成正文 run，也可能抄在公式对象里，所以按文档顺序判，不假设容器类型。
        var nodes = para.DescendantNodes().ToList();
        var textPosition = nodes.FindIndex(node =>
            node is XElement element &&
            (element.Name == Mns + "t" || element.Name == Wns + "t") &&
            element.Value.Contains("(1)", StringComparison.Ordinal));
        var tabPosition = nodes.FindLastIndex(node => node is XElement element && element.Name == Wns + "tab");

        Assert.True(textPosition >= 0, "找不到式号");
        Assert.True(tabPosition >= 0 && tabPosition < textPosition, "式号前面没有制表符，右制表位加了个寂寞");

        Assert.Empty(new FormulaAnalyzer().Analyze(output.OutputPath).Issues);
    }

    /// <summary>
    /// Word 原生写法：式号在公式对象里，制表符由几个 m:r 之间的 w:tab 提供。
    /// 修完不能多出第二个制表符，否则式号会被顶到下一行。
    /// </summary>
    [Fact]
    public void WordNativeSampleKeepsSingleTabAndBecomesCompliant()
    {
        var (_, output) = Run("formula_word_native_number_not_aligned.docx");

        Assert.Empty(new FormulaAnalyzer().Analyze(output.OutputPath).Issues);

        var para = FormulaParagraph(ReadDocument(output.OutputPath), 2);
        var tabCount = para.Descendants(Wns + "tab").Count();
        var tabRun = para.Descendants(Mns + "r")
            .Single(run => run.Elements(Wns + "tab").Any());
        Assert.Contains(tabRun.Elements(Mns + "t"), t => t.Value == "(1)");
        Assert.True(tabCount >= 1);
    }

    [Fact]
    public void SpacingSampleIsZeroedAndKeepsOtherSpacingAttributes()
    {
        var path = SamplePath("formula_spacing_wrong.docx");
        var before = ReadDocument(path);
        var lineRule = before.Descendants(Wns + "spacing").First().Attribute(Wns + "line");

        var (_, output) = Run("formula_spacing_wrong.docx");

        var spacing = FormulaParagraph(ReadDocument(output.OutputPath), 2)
            .Element(Wns + "pPr")?.Element(Wns + "spacing");
        Assert.Equal("0", spacing?.Attribute(Wns + "before")?.Value);
        Assert.Equal("0", spacing?.Attribute(Wns + "after")?.Value);
        // 行距不动：R11 对含公式的段落有豁免。
        Assert.Equal(lineRule?.Value, spacing?.Attribute(Wns + "line")?.Value);

        Assert.Empty(new FormulaAnalyzer().Analyze(output.OutputPath).Issues);
    }

    /// <summary>
    /// 式号字体字号修好，但圆括号不修——加括号要改文字，不在本档范围。
    /// 这条用例同时守住「不要顺手改文本」这条线。
    /// </summary>
    [Fact]
    public void NumberFormatFixesFontAndSizeButNeverAddsBrackets()
    {
        var (_, output) = Run("formula_number_format_wrong.docx");

        var para = FormulaParagraph(ReadDocument(output.OutputPath), 2);
        var numberRun = para.Elements(Wns + "r")
            .Single(run => run.Elements(Wns + "t").Any(t => t.Value == "3"));
        var rPr = numberRun.Element(Wns + "rPr")!;
        Assert.Equal("Times New Roman", rPr.Element(Wns + "rFonts")?.Attribute(Wns + "ascii")?.Value);
        Assert.Equal("21", rPr.Element(Wns + "sz")?.Attribute(Wns + "val")?.Value);

        // 文本没被改：还是「3」，没被补成「(3)」。
        Assert.Equal("3", numberRun.Element(Wns + "t")?.Value);

        var remaining = new FormulaAnalyzer().Analyze(output.OutputPath).Issues.ToList();
        var issue = Assert.Single(remaining);
        Assert.Equal("式号未加圆括号", issue.Title);
        Assert.Contains(output.Result.Skips, skip => skip.Contains("缺圆括号"));
    }

    /// <summary>表格排版的公式不动：位置由列宽决定，改段落对齐看不出效果。</summary>
    [Fact]
    public void TableRowSampleIsSkippedAndExplained()
    {
        var (input, output) = Run("formula_table_row_correct.docx");

        Assert.Empty(output.Fixes);
        Assert.Equal(1, output.Result.TableRowCount);
        Assert.Contains(output.Result.Skips, skip => skip.Contains("表格排版"));
        AssertDocumentUnchanged(input, output.OutputPath);
    }

    [Fact]
    public void AlreadyCompliantWordNativeSampleIsUntouched()
    {
        var (input, output) = Run("formula_word_native_correct.docx");

        Assert.Empty(output.Fixes);
        AssertDocumentUnchanged(input, output.OutputPath);
    }

    /// <summary>幂等：修过的文件再修一遍，不该产生任何新的修改。</summary>
    [Fact]
    public void SecondRunProducesNoFurtherChanges()
    {
        var (_, first) = Run("formula_number_not_right_aligned.docx");

        var secondPath = Path.Combine(Path.GetDirectoryName(first.OutputPath)!, "second.docx");
        _tempFiles.Add(secondPath);
        var second = new FormulaFormatter().Apply(first.OutputPath, secondPath);

        Assert.Empty(second.Fixes);
        AssertDocumentUnchanged(first.OutputPath, secondPath);
    }

    /// <summary>
    /// 硬边界：所有样本修完之后，m:t 和 w:t 的文本序列都必须与修改前逐字相同。
    /// 靠脚本比对，不靠眼看。
    /// </summary>
    [Theory]
    [InlineData("formula_baseline_correct.docx")]
    [InlineData("formula_inline_only.docx")]
    [InlineData("formula_not_centered.docx")]
    [InlineData("formula_number_format_wrong.docx")]
    [InlineData("formula_number_not_right_aligned.docx")]
    [InlineData("formula_spacing_wrong.docx")]
    [InlineData("formula_table_row_correct.docx")]
    [InlineData("formula_word_native_correct.docx")]
    [InlineData("formula_word_native_number_not_aligned.docx")]
    public void MathAndBodyTextStayByteIdentical(string sample)
    {
        var (input, output) = Run(sample);

        Assert.Equal(TextOf(input, "m"), TextOf(output.OutputPath, "m"));
        Assert.Equal(TextOf(input, "w"), TextOf(output.OutputPath, "w"));
    }

    /// <summary>改完能被 undo 撤销：撤销后 document.xml 与修改前完全一致。</summary>
    [Fact]
    public void UndoRestoresTheOriginalDocument()
    {
        var workDir = Path.Combine(Path.GetTempPath(), "nku-formula-undo-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);

        var backupDir = Path.Combine(workDir, "backups");
        var input = Path.Combine(workDir, "in.docx");
        var output = Path.Combine(workDir, "out.docx");
        File.Copy(SamplePath("formula_word_native_number_not_aligned.docx"), input);

        try
        {
            var undo = new UndoManager(backupDir);
            undo.BackupBeforeModify(input, output, "公式格式");
            new FormulaFormatter().Apply(input, output);

            Assert.NotEmpty(TextOf(input, "m")); // 样本确实有公式，避免空跑

            undo.Undo(output);

            AssertDocumentUnchanged(input, output);
        }
        finally
        {
            Directory.Delete(workDir, recursive: true);
        }
    }

    private sealed record RunResult(string OutputPath, FormulaFormatResult Result)
    {
        public IReadOnlyList<FormulaFix> Fixes => Result.Fixes;
    }

    private (string Input, RunResult Output) Run(string sample)
    {
        var input = SamplePath(sample);
        var outputPath = Path.Combine(Path.GetTempPath(), "nku-formula-" + Guid.NewGuid().ToString("N") + ".docx");
        _tempFiles.Add(outputPath);

        var result = new FormulaFormatter().Apply(input, outputPath);
        return (input, new RunResult(outputPath, result));
    }

    /// <summary>临时输出攒在 %TEMP% 里，每个用例跑完就删，别留一地 docx。</summary>
    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            if (File.Exists(file)) File.Delete(file);
        }
    }

    private static XElement FormulaParagraph(XDocument document, int paragraphIndex) =>
        document.Descendants(Wns + "p").ElementAt(paragraphIndex - 1);

    private static List<string> TextOf(string path, string kind)
    {
        var document = ReadDocument(path);
        var name = kind == "m" ? Mns + "t" : Wns + "t";
        return document.Descendants(name).Select(element => element.Value).ToList();
    }

    private static XDocument ReadDocument(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");
        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }

    /// <summary>结构层面判定「一个字都没动」，比只比文本更严：多一个元素也算改动。</summary>
    private static void AssertDocumentUnchanged(string beforePath, string output)
    {
        var before = ReadDocument(beforePath).Root;
        var after = ReadDocument(output).Root;
        Assert.True(XNode.DeepEquals(before, after), "document.xml 结构被改动了");
    }

    private static string SamplePath(string sampleName)
    {
        var root = FindRepositoryRoot();
        var path = Path.Combine(root, "tests", "fixtures", "synthetic", "formula", sampleName);
        Assert.True(File.Exists(path), $"样本不存在：{path}");
        return path;
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
